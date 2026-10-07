# MBABench

MBABench asks an AI agent to do what a junior analyst does: take a business case
and a starting workbook, and build a complete, working financial model in Excel.
It has 101 tasks, four ways of running an agent on them, and an LLM judge that
grades every attempt against a golden solution with a 132-check rubric.

- **Tasks** are on Hugging Face: [namkoong-lab/MBABench](https://huggingface.co/datasets/namkoong-lab/MBABench)
  (starting workbook, golden solution and metadata per task).
- **Agents** reach a model through four surfaces: the vendors' chat products in a browser
  (`gui-agents/`), their add-ins inside Excel Online (`excel-agents/`), raw model APIs driving an
  Excel tool server (`cli-agents/`), and coding agents in a Docker sandbox (`coding-agents/`).
- **Judge** (`judge/`): one LLM conversation per attempt plus deterministic Python checks for the
  Questions-sheet answers and 20 formatting and structure checks.

```text
data/            the downloaded tasks            (scripts/download_dataset.py)
outputs/         attempts and gradings           (written by runs)
config/          config_default.yaml + your config.yaml
house_standards/ the modelling conventions every agent receives with its task
cli-agents/  coding-agents/  gui-agents/  excel-agents/  judge/
```

## Install

You need Python 3.12+, [uv](https://docs.astral.sh/uv/) and LibreOffice (`soffice`, for formula
recalculation). The coding pipeline needs Docker; the GUI and Excel pipelines need Google Chrome
and a paid account with the vendor.

```bash
git clone https://github.com/namkoong-lab/MBABench && cd MBABench
./setup.sh
hf auth login
uv run python scripts/download_dataset.py
```

`setup.sh` creates `.venv`, installs every package and creates `config/config.yaml` from the
defaults. `hf auth login` is needed once because the dataset is private. The download lands in
`data/` and is verified against the dataset's `MANIFEST.json`.

## Configure

`config/config.yaml` (gitignored) holds everything that is yours. Fill in only what you use:

```yaml
keys:
  anthropic_api_key: "${env:ANTHROPIC_API_KEY}"   # or paste the key; CLI + coding pipelines on Claude
  openai_api_key: "${env:OPENAI_API_KEY}"         # CLI + coding pipelines on OpenAI, and the judge
local:
  data_root: "data"       # where download_dataset.py put the tasks
  output_root: "outputs"  # where attempts and gradings go
libreoffice_path: null    # null = find soffice on PATH / in /Applications
```

The GUI and Excel pipelines need no key, only a Chrome signed in to the vendor; their
machine settings (Chrome port, profile, OneDrive folder) go in `<pipeline>/infra/configs/configs.yaml`.

## Run an agent

Every pipeline takes one run config that names the cohort by `agent_model_name` (an entry in the
pipeline's `agent_identities.yaml`, which pins model, effort and settings) and lists the task ids
to run (default: all 101). Runs resume: a task that already has a successful attempt for that
cohort is skipped. Attempts land under `outputs/attempts/<agent_model_name>/`.

**CLI** — a model API plus a local Excel tool server (MCP). Needs the provider key.

```bash
cd cli-agents
uv run excel-agent --batch-config examples/anthropic_claude_fable_5_1_max.yaml --dry-run
uv run excel-agent --batch-config examples/anthropic_claude_fable_5_1_max.yaml
```

**Coding** — Claude Code or Codex inside a Docker sandbox, one container per task. Needs Docker
and `ANTHROPIC_API_KEY` or `OPENAI_API_KEY`.

```bash
cd coding-agents
docker build -t mbabench-coding-agent:v2 docker/
uv run python -m coding_agent.run_sweep --dry-run --config run_configs/claude_code_fable_5_1_max.yaml
uv run python -m coding_agent.run_sweep           --config run_configs/claude_code_fable_5_1_max.yaml
uv run python -m coding_agent.run_sweep --task-ids 1,2 --config run_configs/codex_gpt_6_astra_xhigh.yaml
```

**GUI** — claude.ai or chatgpt.com driven through a real Chrome. Once per provider, launch
the automation Chrome and sign in; the login persists in a profile under `browser_profiles/`
and later runs reuse it (the engine relaunches Chrome itself if it is closed):

```bash
cd gui-agents
scripts/setup_chrome.sh claude
scripts/setup_chrome.sh chatgpt
```

Then run:

```bash
uv run python -m infra.run --dry-run --run-config infra/configs/run_configs/claude_cowork_fable_5_1_max.yaml
uv run python -m infra.run -y        --run-config infra/configs/run_configs/claude_cowork_fable_5_1_max.yaml
```

**Excel** — the Claude or ChatGPT add-in inside Excel Online. Once, launch the automation
Chrome, sign in to Microsoft 365 with the add-in installed, and put the task workbooks in
OneDrive. The engine opens each workbook at `My files / mbabench_tasks / <task_name> / Task /
<workbook>` (the base folder is `onedrive_base_path` in `infra/configs/configs.yaml`). The
quickest way to build that tree is by hand: `--stage` writes it locally under
`onedrive_staging/`, you drag the task folders into `mbabench_tasks` in OneDrive web in one go,
and `--verify` walks the result. Without `--stage` the script uploads through the browser itself,
which is slower and exposed to OneDrive's UI changes.

```bash
cd excel-agents
scripts/setup_chrome.sh
uv run python scripts/provision_onedrive.py --stage
uv run python scripts/provision_onedrive.py --verify
```

Then run:

```bash
uv run python -m infra.run --dry-run --run-config infra/configs/run_configs/claude_excel_fable_5_1.yaml
uv run python -m infra.run -y        --run-config infra/configs/run_configs/claude_excel_fable_5_1.yaml
```

All three browser setups default to Chrome's debugging port 9222 with their own profile, so
run one signed-in Chrome at a time, or give the others another `browser.cdp_port` in the
pipeline's gitignored `infra/configs/configs.yaml`.

`--dry-run` resolves the task, prompt, attachments and output path without touching a model or
a browser. Each example config has a header comment on what to change.

**Try one task first.** A full cohort is 101 attempts; start with one. CLI: set `task_ids: [1]`
in the batch config. Coding: `--task-ids 1`. GUI and Excel: `--task-id 1`. Then grade it with
`grade.py --all` (below) and look at `outputs/gradings/<id>/scores.json`.

## Configure a run

Two files decide what an attempt is. The **run config** you pass on the command line says which
cohort runs, on which tasks, with which prompt and limits. The **identity registry** entry that
the cohort label points to says which model, at which effort, over which route (next section). So
to change the model or the effort you pick or add a registry label; everything else is in the run
config. Each example config documents its keys in a header comment; the ones you are likely to touch:

| pipeline | keys |
|---|---|
| CLI (`examples/*.yaml`) | `agent_model_name`; `task_ids: [...]` or `tasks: {first, last}` (omit = all); `prompt_version` (default `v16`, the leaderboard prompt); `max_iterations` (40 model calls per task); `api_timeout_seconds`; `skip_if_attempted` |
| Coding (`run_configs/*.yaml`) | `agent_model_name`; `tasks: {first, last}`; `template_version` (default `v13`); `sandbox.image`; `--task-ids`, `--workers`, `--redo` on the command line |
| GUI (`infra/configs/run_configs/*.yaml`) | `provider.kind` (`claude` or `chatgpt`) and the `claude_web` / `chatgpt_web` block (`mode`, `model`, `effort`, exactly as the product's UI names them); the label is derived from that block and must match a registered combination; `prompt_version` (205); `source.filters.task_ids` |
| Excel (`infra/configs/run_configs/*.yaml`) | `agent_model_name`; `prompt_version` (205); `source.filters.task_ids`; per-task time caps under the add-in's block |
| Judge (command line) | `--model` picks the grader label; `--workers`; see "Grade attempts" |

Prompt versions are registered text: the CLI and coding pipelines keep theirs under
`prompts/`, the GUI and Excel pipelines under `tasks_configs/prompts/`, each with a registry that
maps the version number to the files sent. Every attempt row records the label and the prompt
version it ran with, so results stay comparable.

## Models and routes

A cohort label (`agent_model_name`) is an entry in the pipeline's identity registry, which pins
the model, the effort and the API route. The shipped entries cover:

| pipeline | registry | routes |
|---|---|---|
| CLI | `cli-agents/excel_cli_agent/agent_identities.yaml` | Anthropic direct, OpenAI direct, TensorBlock Forge; OpenRouter works with an entry whose `base_url` points at it |
| Coding | `coding-agents/coding_agent/agent_identities.yaml` | Claude Code on Anthropic direct; Codex on OpenAI direct or through TensorBlock Forge |
| GUI, Excel | `gui-agents/infra/configs/agent_identity.py`, `excel-agents/agent_identities.yaml` | the vendor's consumer product; no API key |
| Judge | `judge/judge_identities.yaml` | OpenAI, Anthropic, Gemini, OpenRouter, TensorBlock graders |

To run another model or route, append an entry and set the matching key in `config.yaml`
(`anthropic_api_key`, `openai_api_key`, `openrouter_api_key`, `gemini_api_key`, `forge_api_key`).
A CLI entry, for example:

```yaml
- agent_model_name: openpyxl_anthropic/claude-fable-5-1-max
  model: claude-fable-5-1
  reasoning_effort: max
  thinking_budget_tokens: null
  max_completion_tokens: 128000
  base_url: https://api.anthropic.com
  fresh_context_mode: true
  enhanced_excel_context: true
  recent_history_count: 3
```

Labels are append-only: a label's settings never change once attempts have been recorded under it,
so a new configuration gets a new label.

## Grade attempts

The judge reads every `outputs/attempts/**/task_attempts.jsonl`, stages each attempt with its
task's golden solution and rubric annotation, and writes `outputs/gradings/gradings.jsonl` plus one
folder of evidence per grading. The default grader is `openai/gpt-5.6-sol` (about $3.5 and 2
minutes per attempt); `judge/judge_identities.yaml` lists the others.

```bash
uv run python judge/main_scripts/grade.py --all --dry-run
uv run python judge/main_scripts/grade.py --all
uv run python judge/main_scripts/grade.py --agent-model-name <label> --workers 4
uv run python judge/main_scripts/grade.py --attempt-ids <id> --regrade
```

`--dry-run` lists what would be graded. `--all` grades every attempt that has no grading yet;
`--regrade` grades again.

To grade one workbook outside the benchmark flow, assemble a folder with `ai_attempt.xlsx`,
`solution/<golden>.xlsx` and optionally `starting/<workbook>.xlsx`, and run
`uv run python judge/main_scripts/judge.py -f <folder>` (set `JUDGE_SKIP_SUITABILITY=1` for a
task outside the pool). Results land in `<folder>/judge_results/`.

A grading records a verdict per rubric check, a weighted score per category and a total out of
100 (`scores.json`), the Questions-sheet answer check, the deterministic checks, and the judge's
full conversation.

## Troubleshooting

- **Effort values are per model and per endpoint.** A registry entry pins an effort, but the
  provider decides what it accepts, and the answer can differ between the chat and the Responses
  endpoint or once tools are attached. The symptom is a bare 400 such as "provider rejected the
  request". Probe the value with a one-line request before running a cohort, and read the
  recorded upstream request in the attempt's `trajectory.jsonl.gz` to see exactly what was sent.
  Codex only sends its file tools for models described in `coding-agents/docker/codex_model_catalog.json`;
  a model without an entry falls back to a mode the relay cannot carry and the agent gives up at once.
- **UI drift (GUI and Excel).** The vendors rename models and move controls. The engines match
  the product's own labels (`claude_web.model`, `ui_model_label`, the model pill) and refuse to send
  when the screen does not show what the identity claims, so a stuck run usually means the UI changed
  or the account is at a usage cap or signed out. Probe the live UI, then add a registry entry with
  the labels it shows now. `--dry-run` cannot catch any of this; watch the first task of a cohort.
- **Judge.** LibreOffice must be on PATH or set as `libreoffice_path`. A workbook saved without
  cached values (anything written by openpyxl) needs `--run-calculation`, or the judge refuses it as
  ungraded formulas. A task outside the benchmark pool has no rubric-suitability annotation; grade it
  with `JUDGE_SKIP_SUITABILITY=1`. Gradings skip attempts already graded by the same grader; pass
  `--regrade` to grade again.

## Output layout

```text
outputs/
  attempts/<agent_model_name>/
    task_attempts.jsonl                   one row per attempt: task_id, files, timing, cost, agent_failed, ...
    task_id=<N>/<YYYYmmdd_HHMMSS>/        the delivered workbook first, then prompts, transcript, logs
  gradings/
    gradings.jsonl                        one row per grading: attempt_id, grader_model, scores, ...
    <grading_id>/                         scores.json, ai_judgement.json, det_checks.json, extracted CSVs, trajectory
```

## Tasks

Each `data/tasks/task_id=<N>/task.json` carries `task_name`, `difficulty` (Easy to Hard),
`difficulty_score`, `model_type`, a one-line `description`, `estimated_solve_time_hours`, and the
`starting_files` / `solution_files` paths. `data/tasks.jsonl` has the same rows in one file.
