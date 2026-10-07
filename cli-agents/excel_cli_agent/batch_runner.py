#!/usr/bin/env python3
"""
Batch Runner for the Excel CLI Agent.

    excel-agent --batch-config <yaml> [--dry-run]

One attempt per task, sequentially:

  resolve tasks from <data_root>/tasks/task_id=<N>/task.json
    -> skip tasks this cohort already has a non-failed row for
    -> stage the task's starting files (+ the prompt version's attachments)
       into a fresh workspace
    -> run the agent (openpyxl MCP server + model) in that workspace
    -> copy the attempt's files to <output_root>/attempts/<agent_model_name>/task_id=<N>/<ts>/
       and append one row to <output_root>/attempts/<agent_model_name>/task_attempts.jsonl

The batch config names the cohort (agent_model_name — the entry in
agent_identities.yaml supplies every model setting), the prompt set
(prompt_version) and the run limits (max_iterations, api_timeout_seconds);
task selection is task_ids / tasks, default every task present. No database,
no object store: repo_config.data_root() / output_root() are the only two
locations involved.
"""
import json
import os
import shutil
import tempfile
import time
import traceback
import yaml
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from .agent_identity import resolve_agent_identity
from .local_sink import LocalAttemptSink, iso
from .local_source import LocalTask, LocalTaskSource
from .mcp_client import ExcelMCPClient
from .models_config import (DEFAULT_MAX_COMPLETION_TOKENS, DEFAULT_MAX_ITERATIONS, resolve_api_timeout,
                            resolve_stall_timeout, uses_gemini_tool_calls)
from .prompt_versions import (
    DEFAULT_PROMPT_VERSION, PROMPT_VERSIONS, PROMPTS_DIR, attachment_names_for, attachments_for,
    parse_prompt_version, system_prompt_file, template_file,
)
from .repo_config import (
    attachment_extra_configs, data_root, describe_local_target, output_root, resolve_attachments,
    to_output_relative,
)
from .task_executor import ExcelTaskExecutor, TaskStatus

# Per-batch working files (workspaces, summary.md, aggregated_metrics.json):
# cli-agents/batch_logs/batch_<ts>/ — gitignored. The durable record of every
# attempt is the output root, not this directory.
BATCH_LOGS_ROOT: Path = Path(__file__).resolve().parents[1] / "batch_logs"

# Config keys from earlier layouts. Those mapping to a value are accepted when
# they carry that value (the only one this runner implements) and refused
# otherwise; those mapping to None are refused outright with the message.
LEGACY_KEYS: Dict[str, Optional[str]] = {
    "auto_mode": None,          # accepted, ignored (the only mode)
    "benchmark": "v2",          # the benchmark is MBABench, always
    "source": "local",          # tasks come from <data_root>
    "sink": "local",            # attempts go to <output_root>
}
REMOVED_KEYS = {
    "local_mode": "tasks are read from <data_root>/tasks/ (scripts/download_dataset.py); "
                  "there is no folder mode",
    "workspaces": "tasks are read from <data_root>/tasks/; select them with task_ids / tasks",
    "results_dir": "attempts land under <output_root>/attempts/<agent_model_name>/ "
                   "(local.output_root in config/config.yaml or MBABENCH_OUTPUT_ROOT)",
    "task_filter": "select tasks with task_ids / tasks, or omit both for every task present",
    "max_trials": "one attempt per task; a relaunch skips tasks that already have a "
                  "non-failed row (skip_if_attempted: false to re-run them)",
    "trials_since": "see max_trials",
    "skip_if_succeeded": "renamed skip_if_attempted (default true)",
    "task_template": "the prompt version selects the template",
    "task_type": "one shared task template per prompt version",
    "agent_folder": "name the cohort with agent_model_name (agent_identities.yaml)",
}


@dataclass
class WorkspaceConfig:
    """Configuration for a single workspace"""
    path: str
    detected_pdf_files: List[str] = field(default_factory=list)
    detected_excel_files: List[str] = field(default_factory=list)
    # .md files (house standards); embedded verbatim in the model context.
    detected_text_files: List[str] = field(default_factory=list)


@dataclass
class WorkspaceResult:
    """Result of processing a single workspace"""
    workspace_path: str
    status: str  # "success", "failed", "error"
    pdf_files: List[str]
    excel_files: List[str]
    task_id: Optional[str]
    iterations: int
    total_tokens: int
    cost_usd: float
    error_message: Optional[str]
    duration_seconds: float
    final_result: Optional[str]
    start_time: Optional[float] = None   # epoch timestamp
    end_time: Optional[float] = None     # epoch timestamp
    # True if any iteration ran with reduced context (sheet summarization /
    # PDF truncation); None when execution errored before we could know.
    context_reduced: Optional[bool] = None


@dataclass
class BatchResult:
    """Aggregated results from batch processing"""
    batch_name: str
    total_workspaces: int
    successful: int
    failed: int
    workspace_results: List[WorkspaceResult]
    total_duration_seconds: float
    aggregated_tokens: int
    aggregated_iterations: int
    aggregated_cost_usd: float


@dataclass
class TaskInfo:
    """One benchmark task as the runner sees it."""
    task_id: int
    task_name: str
    task_starting_files: List[str] = field(default_factory=list)  # absolute local paths

    @classmethod
    def from_local(cls, task: LocalTask) -> "TaskInfo":
        return cls(task_id=task.task_id, task_name=task.task_name,
                   task_starting_files=list(task.starting_files))


class BatchRunner:
    """Batch execution engine for the Excel CLI Agent (synchronous)."""

    def __init__(self, config_path: str, server_path: str, api_key: str,
                 custom_reasoning: bool = False, enable_langfuse: bool = False):
        self.config_path = Path(config_path)
        self.server_path = server_path
        self.api_key = api_key
        self.custom_reasoning = custom_reasoning
        self.enable_langfuse = enable_langfuse
        self.config: Optional[Dict[str, Any]] = None
        self.batch_logs_dir: Optional[Path] = None
        # Provenance from _verify_recalc_engine(): which formula recalc engine
        # the MCP server runs ({"engine": "libreoffice"|"fallback", ...}).
        self._recalc_engine_info: Optional[Dict[str, Any]] = None
        # Resolved by load_config.
        self._identity = None
        self._prompt_version_label: str = DEFAULT_PROMPT_VERSION
        self.prompt_version: int = 0
        self.system_prompt_path: Path = PROMPTS_DIR / PROMPT_VERSIONS[DEFAULT_PROMPT_VERSION]["system"]
        self.template_path: Path = PROMPTS_DIR / PROMPT_VERSIONS[DEFAULT_PROMPT_VERSION]["template"]
        # Absolute paths of the prompt version's attachments (house
        # standards); copied into every workspace.
        self._attachments: List[Path] = []
        self._attachment_names: Dict[str, str] = {}
        self._source: Optional[LocalTaskSource] = None
        self._sink: Optional[LocalAttemptSink] = None
        self._prompt_files: Optional[List[str]] = None

    # ------------------------------------------------------------------
    # configuration
    # ------------------------------------------------------------------

    def load_config(self) -> Dict[str, Any]:
        """Load and validate the batch configuration."""
        print(f"📋 Loading batch configuration from {self.config_path}")

        with open(self.config_path, 'r') as f:
            config = yaml.safe_load(f) or {}

        for key in ('batch_name', 'agent_model_name'):
            if key not in config:
                raise ValueError(f"Missing required field in config: {key}")

        for key, note in REMOVED_KEYS.items():
            if key in config:
                raise ValueError(f"Config key `{key}` is no longer supported: {note}")
        for key, only in LEGACY_KEYS.items():
            if key in config and only is not None and str(config[key]).lower() != only:
                raise ValueError(
                    f"Config key `{key}` can only be '{only}' (it may also be omitted): "
                    f"the runner reads tasks from <data_root> and writes attempts to "
                    f"<output_root>, nothing else."
                )
            config.pop(key, None)

        config.setdefault('verbose', False)
        config.setdefault('max_iterations', DEFAULT_MAX_ITERATIONS)
        config.setdefault('snapshot_iterations', False)
        config.setdefault('skip_if_attempted', True)
        config.setdefault('cleanup_workspace', True)

        # The config names its cohort (agent_model_name) and NOTHING else
        # about the model: the agent_identities.yaml entry for that label
        # supplies model, reasoning_effort, thinking_budget_tokens,
        # max_completion_tokens, base_url, fresh_context_mode,
        # enhanced_excel_context and recent_history_count. A config that sets
        # any of them refuses to run (resolve_agent_identity), so rows under
        # one label cannot have run with different settings.
        identity = resolve_agent_identity(config)
        config.update(identity.settings())
        self._identity = identity

        prompt_ver = str(config.get('prompt_version', DEFAULT_PROMPT_VERSION))
        if prompt_ver not in PROMPT_VERSIONS:
            raise ValueError(f"Unknown prompt_version '{prompt_ver}'. Available: {list(PROMPT_VERSIONS.keys())}")
        self._prompt_version_label = prompt_ver
        config['prompt_version'] = prompt_ver
        # The set's system prompt - or, for Gemini 3.8 Flash alone, its
        # function-call variant (prompt_versions.MODEL_SYSTEM_PROMPT_VARIANTS);
        # parse_prompt_version still records the set's version for it.
        self.system_prompt_path = PROMPTS_DIR / system_prompt_file(
            prompt_ver, identity.model, require_variant=uses_gemini_tool_calls(identity.model))
        self.template_path = PROMPTS_DIR / template_file(prompt_ver)
        for p in (self.system_prompt_path, self.template_path):
            if not p.is_file():
                raise FileNotFoundError(f"Prompt file not found: {p}")
        self.prompt_version = parse_prompt_version(self.system_prompt_path, self.template_path)
        config['system_prompt_path'] = str(self.system_prompt_path)
        config['task_template'] = self.template_path.read_text(encoding='utf-8')

        # The prompt version — never the config — names the files shipped
        # with every workspace (house standards). Resolved now so a missing
        # file fails the batch before any task is claimed.
        self._attachments = resolve_attachments(attachments_for(prompt_ver))
        self._attachment_names = attachment_names_for(prompt_ver)

        # Where tasks come from and where attempts go. Checked here, before
        # any task is claimed: a missing download is a setup mistake, not a
        # per-task one.
        self._source = LocalTaskSource(data_root())
        self._source.load_all()
        self._sink = LocalAttemptSink(output_root(), config['agent_model_name'])

        self.config = config

        print(f"✅ Configuration loaded: {config['batch_name']}")
        print(f"   Agent model name: {config['agent_model_name']} (agent_identities.yaml)")
        print(f"   Pinned by identity: {identity.settings()}")
        print(f"   Local roots: {describe_local_target()}")
        print(f"   Tasks: {self._source.tasks_dir} ({len(self._source.load_all())} present)")
        print(f"   Attempts: {self._sink.cohort_dir}")
        print(f"   Prompt version: {prompt_ver} (recorded as {self.prompt_version})")
        print(f"   System prompt: {self.system_prompt_path.name}; template: {self.template_path.name}")
        print(f"   Attachments: {[self._delivered_name(p) for p in self._attachments] or 'none'}")
        print(f"   Max iterations: {config['max_iterations']}")
        print(f"   API timeout: {resolve_api_timeout(config.get('reasoning_effort'), config.get('api_timeout_seconds'))}s per call")
        print(f"   Skip if attempted: {config['skip_if_attempted']}")
        print(f"   Workspace base: {config.get('workspace_base_dir') or f'{BATCH_LOGS_ROOT}/batch_<ts>/workspaces (default)'}")

        return config

    # ------------------------------------------------------------------
    # task selection
    # ------------------------------------------------------------------

    def resolve_tasks(self) -> List[TaskInfo]:
        """The tasks this batch covers, in the order it will run them.

        `task_ids: [..]` keeps the given order; `tasks:` is a list of task
        names or a `{first: N, last: M}` id range; neither = every task
        present, ascending id. `skip_task_ids` drops ids from any of them.
        """
        source = self._source
        result: List[TaskInfo] = []

        if 'task_ids' in self.config and self.config['task_ids'] is not None:
            task_ids = list(self.config['task_ids'])
            print(f"\n🔍 Resolving {len(task_ids)} explicit task id(s)...")
            for tid in task_ids:
                task = source.get(tid)
                if task is None:
                    print(f"  ❌ id {tid} -> NOT FOUND under {source.tasks_dir}")
                    continue
                result.append(TaskInfo.from_local(task))
                print(f"  ✅ id {tid} -> {task.task_name}")

        elif isinstance(self.config.get('tasks'), dict):
            rng = self.config['tasks']
            first = int(rng.get('first', 1))
            last = int(rng.get('last', max(t.task_id for t in source.load_all())))
            print(f"\n🔍 Resolving task ids {first}-{last}...")
            result = [TaskInfo.from_local(t) for t in source.load_all() if first <= t.task_id <= last]

        elif self.config.get('tasks'):
            names = list(self.config['tasks'])
            print(f"\n🔍 Resolving {len(names)} task name(s)...")
            for name in names:
                task = source.find_by_name(str(name))
                if task is None:
                    print(f"  ❌ {name} -> NOT FOUND under {source.tasks_dir}")
                    continue
                result.append(TaskInfo.from_local(task))
                print(f"  ✅ {name} -> id {task.task_id}")

        else:
            print(f"\n🔍 Every task under {source.tasks_dir}...")
            result = [TaskInfo.from_local(t) for t in source.load_all()]

        skip_ids = set(self.config.get('skip_task_ids') or [])
        if skip_ids:
            dropped = sorted({t.task_id for t in result} & skip_ids)
            if dropped:
                print(f"⏭️  Skipping {len(dropped)} task id(s) via skip_task_ids: {dropped}")
            result = [t for t in result if t.task_id not in skip_ids]

        print(f"📋 Total tasks selected: {len(result)}")
        return result

    def should_skip(self, task_info: TaskInfo) -> bool:
        """True if this cohort already recorded a non-failed attempt of the
        task under this prompt version (the resume rule)."""
        if not self.config.get('skip_if_attempted', True):
            return False
        if self._sink.has_attempt(task_info.task_id, self.prompt_version):
            print(f"  ⏭️  {task_info.task_name}: already attempted (row in "
                  f"{to_output_relative(self._sink.jsonl_path)}), skipping")
            return True
        return False

    # ------------------------------------------------------------------
    # workspace
    # ------------------------------------------------------------------

    def setup_workspace(self, task_info: TaskInfo) -> str:
        """Create the workspace and stage the task's starting files into it."""
        base_override = self.config.get('workspace_base_dir')
        if base_override:
            base_dir = Path(base_override).expanduser()
        elif self.batch_logs_dir is not None:
            base_dir = Path(self.batch_logs_dir) / "workspaces"
        else:
            raise RuntimeError(
                "No workspace location: set workspace_base_dir in the config, "
                "or run via run_batch so the batch_logs directory exists."
            )
        folder_name = task_info.task_name.replace(' ', '_')
        run_id = f"{int(time.time())}_{os.getpid()}"
        workspace = base_dir / f"{folder_name}_{run_id}"
        workspace.mkdir(parents=True, exist_ok=True)
        print(f"  📁 Workspace: {workspace}")

        # A task with no starting files would run the agent with no task
        # content at all (the empty-context defect) — refuse it.
        if not task_info.task_starting_files:
            raise RuntimeError(
                f"Task '{task_info.task_name}' has no starting_files; "
                "refusing to run the agent with an empty workspace"
            )

        for src in task_info.task_starting_files:
            src_path = Path(src)
            local_path = workspace / src_path.name
            try:
                shutil.copy2(src_path, local_path)
            except Exception as e:
                raise RuntimeError(
                    f"Failed to copy {src_path} for '{task_info.task_name}': {e}. "
                    "Aborting this task — running without starting files would "
                    "produce an invalid (empty-context) attempt."
                ) from e
            size = local_path.stat().st_size
            if size == 0:
                raise RuntimeError(
                    f"Copied zero-byte file {src_path.name} for '{task_info.task_name}'; "
                    "refusing to run on an empty starting file"
                )
            print(f"  ✅ Staged: {src_path.name} ({size:,} bytes)")

        # The prompt version's attachments (house standards) ride alongside
        # the starting files, under the bare name the prompt directive uses.
        self._copy_attachments(workspace)
        return str(workspace)

    def detect_workspace_files(self, workspace_path: str) -> WorkspaceConfig:
        """Auto-detect the .xlsx, .pdf and .md context files in a workspace."""
        workspace = Path(workspace_path)
        if not workspace.exists():
            raise ValueError(f"Workspace does not exist: {workspace_path}")

        # Every .xlsx except solution.xlsx is Excel context.
        xlsx_files = [f for f in workspace.glob("*.xlsx") if f.name.lower() != "solution.xlsx"]
        excel_context_files = [str(f.name) for f in xlsx_files]

        # Bare names: consumers join these onto the workspace path, so a
        # prefixed path would double the workspace segment and silently
        # resolve to a missing file (the empty-PDF-context defect).
        pdf_files = [str(f.name) for f in workspace.glob("*.pdf")]
        text_files = [str(f.name) for f in workspace.glob("*.md")]

        return WorkspaceConfig(
            path=workspace_path,
            detected_pdf_files=pdf_files,
            detected_excel_files=excel_context_files,
            detected_text_files=text_files,
        )

    def estimate_context_tokens(self, workspace_path: str, pdf_files: List[str], excel_files: List[str]) -> int:
        """Rough token estimate of the PDF and Excel context, for the log."""
        workspace = Path(workspace_path)
        total = 0
        for pdf_file in pdf_files:
            p = workspace / pdf_file
            if p.exists():
                total += p.stat().st_size // 4
        for excel_file in excel_files:
            p = workspace / excel_file
            if p.exists():
                total += p.stat().st_size // 6
        return total

    def cleanup_workspace(self, workspace_path: str):
        """Delete the workspace once the attempt is recorded."""
        if self.config.get('cleanup_workspace', True) is False:
            print(f"  📂 Workspace preserved: {workspace_path}")
            return
        try:
            shutil.rmtree(workspace_path)
            print(f"  🧹 Cleaned up workspace: {workspace_path}")
        except Exception as e:
            print(f"  ⚠️  Cleanup failed for {workspace_path}: {e}")

    # ------------------------------------------------------------------
    # recalc engine preflight
    # ------------------------------------------------------------------

    def _server_extra_args(self) -> List[str]:
        """Extra argv for the MCP server subprocess.

        allow_recalc_fallback: true lets the server run with the degraded
        _eval_formula engine when LibreOffice is unavailable. By default an
        unavailable engine makes the server exit — and so the batch fail
        loudly — instead of silently producing fallback-engine attempts.
        """
        if self.config and self.config.get('allow_recalc_fallback'):
            return ["--allow-recalc-fallback"]
        return []

    def _verify_recalc_engine(self):
        """Abort the batch if the recalc engine can't start; record provenance.

        Spawns one throwaway MCP server (exactly as each workspace will) and
        asks it which engine it runs. With the default strict server, a
        machine without LibreOffice fails here — before any task is claimed
        — instead of mid-batch. The answer is recorded on every attempt row
        (extra_configs.recalc_engine).
        """
        probe_dir = tempfile.mkdtemp(prefix="recalc_probe_")
        client = ExcelMCPClient(self.server_path, probe_dir,
                                server_args=self._server_extra_args())
        try:
            client.connect()
            result = client.call_tool("get_recalc_engine_info", {})
            if not result.get("success") or not isinstance(result.get("result"), dict):
                raise RuntimeError(f"get_recalc_engine_info returned {result}")
            self._recalc_engine_info = result["result"]
        except Exception as e:
            raise RuntimeError(
                "Recalc engine preflight failed: the Excel MCP server did not "
                f"come up ({e}). Most likely LibreOffice is missing — install "
                "it (Linux: apt-get install libreoffice-calc; macOS: install "
                "LibreOffice.app) or set libreoffice_path in "
                "<MBABench>/config/config.yaml. To deliberately run with the "
                "degraded _eval_formula fallback, set allow_recalc_fallback: "
                "true in the batch config."
            ) from e
        finally:
            try:
                client.disconnect()
            except Exception:
                pass
            shutil.rmtree(probe_dir, ignore_errors=True)

        if self._recalc_engine_info.get("engine") == "libreoffice":
            version = self._recalc_engine_info.get("soffice_version") or "version unknown"
            print(f"  ✅ Recalc engine verified: libreoffice ({version})")
        else:
            print("  ⚠️  Recalc engine: _eval_formula fallback "
                  "(allow_recalc_fallback set) — attempts will be recorded "
                  "with recalc_engine=fallback.")

    # ------------------------------------------------------------------
    # attachments + extra_configs
    # ------------------------------------------------------------------

    def _delivered_name(self, src: Path) -> str:
        """The filename an attachment lands under in the workspace: the
        version's `attachment_names` mapping, else the source's own name."""
        return self._attachment_names.get(src.name, src.name)

    def _copy_attachments(self, workspace: Path) -> None:
        """Copy the prompt version's attachments into the workspace under
        the name the prompt directive uses (v15+: HOUSE_STANDARDS.md).

        An attachment that fails to land would run the agent against a
        prompt promising text it never saw, so a missing/empty source or
        copy is fatal for the task.
        """
        for src in self._attachments:
            if not src.is_file() or src.stat().st_size == 0:
                raise RuntimeError(
                    f"Prompt attachment missing or empty: {src}; refusing to "
                    "run the agent without the text its prompt version promises"
                )
            dst = workspace / self._delivered_name(src)
            shutil.copy2(src, dst)
            if not dst.is_file() or dst.stat().st_size == 0:
                raise RuntimeError(f"Failed to copy prompt attachment {src.name} into {workspace}")
            print(f"  📎 Attached: {dst.name} ({dst.stat().st_size:,} bytes)")

    def _attachment_extra_configs(self) -> Dict[str, Any]:
        """The attachment-provenance keys merged into extra_configs per attempt."""
        return attachment_extra_configs(self._attachments, self._attachment_names)

    def _response_contract_extra_configs(self) -> Dict[str, Any]:
        """How the model was asked to answer, merged into extra_configs per
        attempt. Every model but Gemini 3.8 Flash answers in JSON text with
        the prompt set's own system prompt (json_actions); that one model
        (models_config.GEMINI_TOOL_CALL_MODELS) answers with native function
        calls and runs the set's variant system prompt, recorded here by file
        name while prompt_version stays the set's (1609 for v16)."""
        native = uses_gemini_tool_calls(self._identity.model)
        cfg = {"response_contract": "native_tools" if native else "json_actions"}
        set_file = PROMPT_VERSIONS[self._prompt_version_label]["system"]
        if self.system_prompt_path.name != set_file:
            cfg["system_prompt_file"] = self.system_prompt_path.name   # a variant: name it
        return cfg

    def _run_limit_extra_configs(self) -> Dict[str, Any]:
        """The two run limits, merged into extra_configs per attempt: how many
        iterations (model calls) the agent was allowed and how long one call
        could take."""
        limits = {
            "max_iterations": int(self.config.get("max_iterations", DEFAULT_MAX_ITERATIONS)),
            "api_timeout_seconds": resolve_api_timeout(self.config.get("reasoning_effort"),
                                                       self.config.get("api_timeout_seconds")),
        }
        # Forge rows only: the silence after which a try is cut and retried
        # inside that per-call budget.
        stall = resolve_stall_timeout(self.config.get("base_url"), self.config.get("model"))
        if stall:
            limits["stream_stall_seconds"] = stall
        return limits

    def _recalc_extra_configs(self) -> Dict[str, Any]:
        """The recalc-provenance keys merged into extra_configs per attempt."""
        if not self._recalc_engine_info:
            return {}
        out: Dict[str, Any] = {"recalc_engine": self._recalc_engine_info.get("engine")}
        if self._recalc_engine_info.get("soffice_version"):
            out["libreoffice_version"] = self._recalc_engine_info["soffice_version"]
        return out

    def _extra_configs(self) -> Dict[str, Any]:
        """Everything an attempt row records about how it ran."""
        cfg = dict(self._identity.extra_configs())
        cfg.update(self._run_limit_extra_configs())
        cfg.update(self._recalc_extra_configs())
        cfg.update(self._attachment_extra_configs())
        cfg.update(self._response_contract_extra_configs())
        return cfg

    # ------------------------------------------------------------------
    # one workspace = one agent run
    # ------------------------------------------------------------------

    def process_workspace(self, workspace_config: WorkspaceConfig) -> WorkspaceResult:
        """Run the agent on one workspace (synchronous)."""
        workspace_path = workspace_config.path
        start_time = time.time()

        print(f"\n{'='*80}")
        print(f"🚀 Processing workspace: {workspace_path}")
        print(f"{'='*80}")

        result = WorkspaceResult(
            workspace_path=workspace_path,
            status="error",
            pdf_files=workspace_config.detected_pdf_files,
            excel_files=workspace_config.detected_excel_files,
            task_id=None,
            iterations=0,
            total_tokens=0,
            cost_usd=0.0,
            error_message=None,
            duration_seconds=0,
            final_result=None
        )

        try:
            print(f"📄 PDF files: {len(workspace_config.detected_pdf_files)}")
            for pdf in workspace_config.detected_pdf_files:
                print(f"   - {Path(pdf).name}")
            print(f"📊 Excel files for context: {len(workspace_config.detected_excel_files)}")
            for excel in workspace_config.detected_excel_files:
                print(f"   - {excel}")
            print(f"📎 Text files for context: {len(workspace_config.detected_text_files)}")
            for text in workspace_config.detected_text_files:
                print(f"   - {text}")

            estimated_tokens = self.estimate_context_tokens(
                workspace_path,
                workspace_config.detected_pdf_files,
                workspace_config.detected_excel_files
            )
            print(f"📏 Estimated context tokens: {estimated_tokens:,}")

            task_description = self.config['task_template']
            print(f"📝 Task: {task_description[:100]}...")

            excel_client = ExcelMCPClient(self.server_path, workspace_path,
                                          server_args=self._server_extra_args())
            task_executor = ExcelTaskExecutor(
                excel_client,
                self.api_key,
                model=self.config['model'],
                custom_reasoning=self.custom_reasoning,
                fresh_context_mode=self.config.get('fresh_context_mode', False),
                enhanced_excel_context=self.config.get('enhanced_excel_context', True),
                recent_history_count=self.config.get('recent_history_count', 5),
                max_completion_tokens=self.config.get('max_completion_tokens', DEFAULT_MAX_COMPLETION_TOKENS),
                reasoning_effort=self.config.get('reasoning_effort', None),
                api_timeout_seconds=self.config.get('api_timeout_seconds', None),
                base_url=self.config.get('base_url', None),
                thinking_budget_tokens=self.config.get('thinking_budget_tokens', None),
                system_prompt_path=self.config.get('system_prompt_path', None),
            )

            task_executor.set_max_iterations(self.config['max_iterations'])
            task_executor.set_verbose(self.config['verbose'])
            task_executor.snapshot_iterations = self.config['snapshot_iterations']

            excel_client.connect()

            # Register context. A detected file that fails to register means
            # the agent would run without task content — hard-fail instead
            # of producing an invalid (empty-context) attempt.
            if workspace_config.detected_pdf_files:
                add_result = task_executor.add_context_pdfs(workspace_config.detected_pdf_files)
                print(f"✅ Added {len(add_result['added'])} PDFs to context")
                if add_result['not_found']:
                    raise RuntimeError(
                        f"Detected PDFs failed to register as context: "
                        f"{add_result['not_found']} — aborting this task"
                    )

            if workspace_config.detected_excel_files:
                add_result = task_executor.add_context_excels(workspace_config.detected_excel_files)
                print(f"✅ Added {len(add_result['added'])} Excel file(s) to context")
                if add_result['not_found']:
                    raise RuntimeError(
                        f"Detected Excel files failed to register as context: "
                        f"{add_result['not_found']} — aborting this task"
                    )

            if workspace_config.detected_text_files:
                add_result = task_executor.add_context_texts(workspace_config.detected_text_files)
                print(f"✅ Added {len(add_result['added'])} text file(s) to context")
                if add_result['not_found']:
                    raise RuntimeError(
                        f"Detected text files failed to register as context: "
                        f"{add_result['not_found']} — aborting this task"
                    )

            print(f"🧠 Model: {self.config['model']}")
            print(f"⚙️  Max iterations: {self.config['max_iterations']}")

            task_execution = task_executor.execute_task(task_description)
            excel_client.disconnect()

            result.task_id = task_execution.task_id
            result.iterations = task_execution.total_iterations
            result.final_result = task_execution.final_result
            result.context_reduced = task_execution.context_reduced

            if task_execution.status == TaskStatus.COMPLETED:
                result.status = "success"
                print(f"✅ Workspace completed successfully")
            else:
                result.status = "failed"
                result.error_message = task_execution.error or f"Task status: {task_execution.status.value}"
                print(f"❌ Workspace failed: {result.error_message}")

            result.cost_usd = task_execution.total_cost_usd
            result.start_time = task_execution.start_time
            result.end_time = task_execution.end_time

        except Exception as e:
            result.status = "error"
            result.error_message = str(e)
            print(f"💥 Error processing workspace: {str(e)}")

        finally:
            result.duration_seconds = time.time() - start_time
            print(f"⏱️  Duration: {result.duration_seconds:.2f}s")

        return result

    # ------------------------------------------------------------------
    # recording
    # ------------------------------------------------------------------

    def snapshot_prompts(self) -> List[str]:
        """Snapshot this batch's prompts under attempts/<label>/prompts/ once:
        system prompt, task template, then the prompt version's attachments.
        Every row of the batch lists them as prompt_files."""
        if self._prompt_files is not None:
            return self._prompt_files

        prompts_dir = self._sink.cohort_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sources: List[Tuple[Path, str]] = [
            (self.system_prompt_path, self.system_prompt_path.name),
            (self.template_path, self.template_path.name),
        ]
        for attachment in self._attachments:
            sources.append((attachment, self._delivered_name(attachment)))

        stored: List[str] = []
        for src_path, name in sources:
            dest = prompts_dir / f"{timestamp}_{name}"
            shutil.copy2(src_path, dest)
            rel = to_output_relative(dest)
            stored.append(rel)
            print(f"  📝 Prompt snapshot: {rel}")
        self._prompt_files = stored
        return stored

    def _attempt_file_set(self, workspace: Path) -> List[Tuple[Path, str]]:
        """(source, relative destination) pairs for one attempt's folder.

        solution.xlsx is the only workbook recorded (starting files come from
        the task folder; intermediate workbook copies are not part of the
        record), and it sorts first so attempt_files leads with it. Then the
        rest of the workspace (PDFs, the attached standards, agent_logs/
        with the transcript and request log), the exact prompt files, and
        the batch config.
        """
        def _include(p: Path) -> bool:
            if p.name == "solution.xlsx":
                return True
            if p.name.startswith('.'):
                # a save killed mid-write can leave .solution.xlsx.tmp-<pid>
                # behind; dotfiles are never part of an attempt.
                return False
            return not p.name.lower().endswith(('.xlsx', '.xlsm', '.xlsb', '.xls'))

        files = [
            (p, p.relative_to(workspace).as_posix())
            for p in sorted(
                (q for q in workspace.rglob('*') if q.is_file() and _include(q)),
                key=lambda q: (q.name != "solution.xlsx", str(q)),
            )
        ]
        for p in dict.fromkeys([self.system_prompt_path, self.template_path]):
            files.append((p, f"prompts/{p.name}"))
        files.append((self.config_path, f"config/{self.config_path.name}"))
        return files

    def record_result(self, task_info: TaskInfo, workspace_path: str, workspace_result: WorkspaceResult):
        """Copy the attempt's files into the output tree and append its row."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        workspace = Path(workspace_path)

        attempt_files = self._sink.store_files(
            self._attempt_file_set(workspace), task_info.task_id, timestamp)
        print(f"  💾 {len(attempt_files)} file(s) -> {self._sink.attempt_dir(task_info.task_id, timestamp)}")
        if not attempt_files or not attempt_files[0].endswith("solution.xlsx"):
            print("  ⚠️  No solution.xlsx in this attempt — the judge will have nothing to grade")

        start_time_dt = datetime.fromtimestamp(workspace_result.start_time) if workspace_result.start_time else datetime.now()
        end_time_dt = datetime.fromtimestamp(workspace_result.end_time) if workspace_result.end_time else datetime.now()
        time_taken_min = workspace_result.duration_seconds / 60.0 if workspace_result.duration_seconds > 0 else 0.0
        total_cost = workspace_result.cost_usd if workspace_result.cost_usd > 0 else 0.0

        # Hitting the iteration cap is NOT a failure: the workbook was still
        # built and recorded, and it gets judged as-is.
        agent_failed = workspace_result.status != "success"
        agent_failed_reason = workspace_result.error_message if agent_failed else None
        if agent_failed and (agent_failed_reason or "").startswith("Max iterations"):
            agent_failed = False

        now = datetime.now().astimezone()
        row = self._sink.append_row({
            "id": self._sink.new_attempt_id(),
            "task_id": task_info.task_id,
            "agent_model_name": self.config['agent_model_name'],
            "agent_model_type": "api",
            "attempt_files": attempt_files,
            "prompt_files": self.snapshot_prompts(),
            "start_time": iso(start_time_dt),
            "end_time": iso(end_time_dt),
            "time_taken_min": time_taken_min,
            "cost": total_cost,
            "prompt_version": self.prompt_version,
            "agent_failed": agent_failed,
            "agent_failed_reason": agent_failed_reason,
            "deprecated": False,
            "created_at": iso(now),
            "context_reduced": workspace_result.context_reduced,
            "deprecated_reason": None,
            "updated_at": iso(now),
            "extra_configs": self._extra_configs(),
        })
        print(f"  ✅ Recorded attempt (ID: {row['id']}, cost: ${total_cost:.4f}) "
              f"-> {to_output_relative(self._sink.jsonl_path)}")

    # ------------------------------------------------------------------
    # drivers
    # ------------------------------------------------------------------

    def dry_run(self) -> int:
        """Resolve everything a batch would use and print the plan; start no
        MCP server and call no model. Returns a process exit code."""
        self.load_config()
        task_infos = self.resolve_tasks()
        todo = [t for t in task_infos if not self._sink.has_attempt(t.task_id, self.prompt_version)]
        skipped = len(task_infos) - len(todo)
        if not self.config.get('skip_if_attempted', True):
            todo, skipped = task_infos, 0

        print(f"\n{'='*80}")
        print(f"🧪 DRY RUN — {self.config['batch_name']}")
        print(f"{'='*80}")
        for t in task_infos:
            mark = "⏭️ " if t not in todo else "▶ "
            print(f"  {mark} task {t.task_id}: {t.task_name}")
            for f in t.task_starting_files:
                print(f"       - {Path(f).name} ({Path(f).stat().st_size:,} bytes)")
            for a in self._attachments:
                print(f"       + {self._delivered_name(a)} (attachment)")
        print(f"\n  {len(task_infos)} selected, {skipped} already attempted, {len(todo)} to run")
        print(f"  would stage workspaces under: "
              f"{self.config.get('workspace_base_dir') or f'{BATCH_LOGS_ROOT}/batch_<ts>/workspaces'}")
        print(f"  would write: {self._sink.cohort_dir}/task_id=<N>/<YYYYmmdd_HHMMSS>/ "
              f"(+ a row in {self._sink.jsonl_path})")
        print(f"  would run: {self.config['model']} via {self.config.get('base_url')} "
              f"(prompt_version {self.prompt_version}, up to {self.config['max_iterations']} calls/task) — not started")
        print(f"  recalc engine preflight (LibreOffice) skipped in dry run")
        return 0

    def run_batch(self) -> BatchResult:
        """Execute the batch: resolve tasks -> stage -> run -> record."""
        batch_start_time = time.time()
        config = self.load_config()

        # Fail fast if the LibreOffice recalc engine can't start (unless the
        # config sets allow_recalc_fallback), and record which engine this
        # batch runs so every attempt row carries the provenance.
        self._verify_recalc_engine()

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.batch_logs_dir = BATCH_LOGS_ROOT / f"batch_{timestamp}"
        self.batch_logs_dir.mkdir(parents=True, exist_ok=True)
        print(f"\n📂 Batch logs directory: {self.batch_logs_dir}")

        print(f"\n📝 Recording prompt snapshots...")
        self.snapshot_prompts()

        task_infos = self.resolve_tasks()
        if not task_infos:
            print("⚠️  No tasks to process. Exiting.")
            return BatchResult(
                batch_name=config['batch_name'], total_workspaces=0, successful=0, failed=0,
                workspace_results=[], total_duration_seconds=time.time() - batch_start_time,
                aggregated_tokens=0, aggregated_iterations=0, aggregated_cost_usd=0.0,
            )

        workspace_results = []
        skipped_count = 0

        for idx, task_info in enumerate(task_infos):
            print(f"\n{'='*80}")
            print(f"📦 Task {idx + 1}/{len(task_infos)}: {task_info.task_name} (id {task_info.task_id})")
            print(f"{'='*80}")

            if self.should_skip(task_info):
                skipped_count += 1
                continue

            workspace_path = None
            try:
                workspace_path = self.setup_workspace(task_info)
                ws_config = self.detect_workspace_files(workspace_path)
                result = self.process_workspace(ws_config)
                workspace_results.append(result)

                print(f"\n  💾 Recording results...")
                self.record_result(task_info, workspace_path, result)
                self.cleanup_workspace(workspace_path)

            except Exception as e:
                print(f"  💥 Error processing task '{task_info.task_name}': {e}")
                traceback.print_exc()
                workspace_results.append(WorkspaceResult(
                    workspace_path=workspace_path or "unknown",
                    status="error",
                    pdf_files=[],
                    excel_files=[],
                    task_id=None,
                    iterations=0,
                    total_tokens=0,
                    cost_usd=0.0,
                    error_message=str(e),
                    duration_seconds=0,
                    final_result=None,
                ))

        total_duration = time.time() - batch_start_time
        successful = sum(1 for r in workspace_results if r.status == "success")
        batch_result = BatchResult(
            batch_name=config['batch_name'],
            total_workspaces=len(workspace_results),
            successful=successful,
            failed=len(workspace_results) - successful,
            workspace_results=workspace_results,
            total_duration_seconds=total_duration,
            aggregated_tokens=sum(r.total_tokens for r in workspace_results),
            aggregated_iterations=sum(r.iterations for r in workspace_results),
            aggregated_cost_usd=sum(r.cost_usd for r in workspace_results),
        )
        self.generate_reports(batch_result)
        if skipped_count > 0:
            print(f"⏭️  Skipped (already attempted): {skipped_count}")
        return batch_result

    def generate_reports(self, batch_result: BatchResult):
        """Write summary.md + aggregated_metrics.json into the batch logs dir."""
        summary_lines = [
            f"# Batch Execution Summary: {batch_result.batch_name}",
            "",
            f"**Execution Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"**Total Duration**: {batch_result.total_duration_seconds:.2f} seconds",
            f"**Attempts recorded under**: {self._sink.cohort_dir}",
            "",
            "## Overview",
            "",
            f"- **Total Workspaces**: {batch_result.total_workspaces}",
            f"- **Successful**: {batch_result.successful}",
            f"- **Failed**: {batch_result.failed}",
            f"- **Total Iterations**: {batch_result.aggregated_iterations}",
            f"- **Total Tokens**: {batch_result.aggregated_tokens:,}",
            f"- **Total Cost**: ${batch_result.aggregated_cost_usd:.4f}",
            "",
            "## Workspace Results",
            ""
        ]
        for result in batch_result.workspace_results:
            status_emoji = "✅" if result.status == "success" else "❌"
            summary_lines += [
                f"### {status_emoji} {result.workspace_path}",
                "",
                f"- **Status**: {result.status}",
                f"- **PDF Files**: {len(result.pdf_files)}",
                f"- **Excel Context Files**: {len(result.excel_files)}",
                f"- **Iterations**: {result.iterations}",
                f"- **Tokens**: {result.total_tokens:,}",
                f"- **Cost**: ${result.cost_usd:.4f}",
                f"- **Duration**: {result.duration_seconds:.2f}s",
            ]
            if result.final_result:
                summary_lines.append(f"- **Result**: {result.final_result}")
            if result.error_message:
                summary_lines.append(f"- **Error**: {result.error_message}")
            summary_lines.append("")

        (self.batch_logs_dir / "summary.md").write_text("\n".join(summary_lines))

        metrics = {
            "batch_name": batch_result.batch_name,
            "execution_timestamp": datetime.now().isoformat(),
            "attempts_dir": str(self._sink.cohort_dir),
            "total_workspaces": batch_result.total_workspaces,
            "successful": batch_result.successful,
            "failed": batch_result.failed,
            "total_duration_seconds": batch_result.total_duration_seconds,
            "aggregated_tokens": batch_result.aggregated_tokens,
            "aggregated_iterations": batch_result.aggregated_iterations,
            "aggregated_cost_usd": round(batch_result.aggregated_cost_usd, 4),
            "workspace_results": [
                {
                    "workspace_path": r.workspace_path,
                    "status": r.status,
                    "task_id": r.task_id,
                    "pdf_count": len(r.pdf_files),
                    "excel_context_count": len(r.excel_files),
                    "iterations": r.iterations,
                    "tokens": r.total_tokens,
                    "cost_usd": round(r.cost_usd, 4),
                    "duration_seconds": r.duration_seconds,
                    "error": r.error_message
                }
                for r in batch_result.workspace_results
            ]
        }
        (self.batch_logs_dir / "aggregated_metrics.json").write_text(json.dumps(metrics, indent=2))

        print(f"\n{'='*80}")
        print(f"📊 BATCH EXECUTION COMPLETE")
        print(f"{'='*80}")
        print(f"✅ Successful: {batch_result.successful}/{batch_result.total_workspaces}")
        print(f"❌ Failed: {batch_result.failed}/{batch_result.total_workspaces}")
        print(f"📈 Total iterations: {batch_result.aggregated_iterations}")
        print(f"🔢 Total tokens: {batch_result.aggregated_tokens:,}")
        print(f"💰 Total cost: ${batch_result.aggregated_cost_usd:.4f}")
        print(f"⏱️  Total duration: {batch_result.total_duration_seconds:.2f}s")
        print(f"💾 Attempts: {self._sink.cohort_dir}")
        print(f"📂 Reports saved to: {self.batch_logs_dir}")
        print(f"{'='*80}")


def run_batch_from_config(config_path: str, server_path: str, api_key: str,
                          custom_reasoning: bool = False, enable_langfuse: bool = False) -> BatchResult:
    """Entry point for batch execution (synchronous)."""
    runner = BatchRunner(config_path, server_path, api_key, custom_reasoning, enable_langfuse)
    return runner.run_batch()


def dry_run_from_config(config_path: str, server_path: str) -> int:
    """Entry point for `--batch-config ... --dry-run`."""
    runner = BatchRunner(config_path, server_path, api_key="")
    return runner.dry_run()
