"""Prompt assembly.

PROMPT.md = system wrapper + task template + workspace file listing.

Task templates (all task-invariant: one shared file per version):
  v8  — mirror of the GUI prompt with the 132-check rubric embedded
        byte-exact.
  v9  — v8 + the Questions-sheet answer convention (mirror of the v3 GUI
        prompt files; carries the 2026-08 rubric revision, so its rubric
        differs from v8's).
  v10 / v11 — rubric-effect experiment (2026-09-08, frozen: prompt_version
        110/111 have recorded runs): v9 with every rubric-derived passage
        removed (v10); v10 + a pointer to HOUSE_STANDARDS.md, staged into the
        workspace root from prompts/house_standards_v1.md (v11).
  v12 — v9 + the House Standards directive (mirror of the v4 GUI prompt
        files; rubric byte-identical to v9's). The file it names,
        House_Standards_v1.md, is seeded into starting_files/ by the runner
        (config.TEMPLATE_ATTACHMENTS). Superseded as the default by v13.
  v13 — the default (2026-09-10, the 101-task run): the v11 text
        BYTE-IDENTICAL — rubric-free, plus the pointer to HOUSE_STANDARDS.md
        staged into the workspace root — under a new number so the cohort
        never merges with the 111 experiment rows.
  v14 / v15 — prompt ablation (2026-09-23) on the Claude Code / Fable 5.1 max
        identity, v13 staying the master prompt: v14 = the v9 text (the
        rubric added back, NO house standards); v15 = the v10 text (rubric-
        free, no pointer, nothing staged). Neither declares extras or
        attachments.

prompt_version (the attempt row's column) is system*100 + template:
system_prompt_coding_v1 + v8 -> 108; + v9 -> 109; + v10 -> 110; + v11 -> 111;
+ v12 -> 112; + v13 -> 113; + v14 -> 114; + v15 -> 115.
"""
import hashlib
import json
import re
from pathlib import Path

from .config import PROMPTS_DIR, RunConfig
from .task_source import TaskSpec

# v8 embeds the 132-check rubric byte-exact from
# gui-agents/tasks_configs/prompts_v2/step2_build.txt ("== FULL RUBRIC"
# onward). Never edit the rubric section by hand.
V8_RUBRIC_MD5 = "ff2e3d483da45b722ed1a39db07d1c14"
V8_RUBRIC_LEN = 60470
V8_RUBRIC_MARKER = "== FULL RUBRIC"

# v9 (Questions-sheet revision) embeds the 132-check rubric byte-exact from
# gui-agents/tasks_configs/prompts_v3/step2_build.txt. Since the 2026-08
# rubric revision it deliberately DIFFERS from v8's frozen rubric. Never edit
# the rubric section by hand.
V9_RUBRIC_MD5 = "b11e174d69a74d239437c572584041ab"
V9_RUBRIC_LEN = 68761

# v12 (House Standards revision) embeds the rubric byte-exact from
# gui-agents/tasks_configs/prompts_v4/step2_build.txt, whose rubric is
# byte-identical to prompts_v3's — hence the same constants as v9. Kept as
# separate names so a future prompts_v4 rubric edit changes v12's guard
# without touching v9's.
V12_RUBRIC_MD5 = "b11e174d69a74d239437c572584041ab"
V12_RUBRIC_LEN = 68761

# v10 / v11 (rubric-effect experiment 2026-09-08): v9 with every
# rubric-derived passage removed (v10), plus a pointer to the house-standards
# file staged into the workspace root (v11). Whole-file md5 pinned so the
# text recorded under prompt_version 110/111 cannot drift.
SCRUBBED_MD5 = {
    "task_template_shared_v10.txt": "5b607903172352e82f8fbbac986d96ef",
    "task_template_shared_v11.txt": "008713b35953c2d1356dd81e7f51d4e3",
    # v13 = v11 byte-identical; same pin.
    "task_template_shared_v13.txt": "008713b35953c2d1356dd81e7f51d4e3",
    # v15 = v10 byte-identical (ablation arm "minus the house standards");
    # same pin as v10, same rubric-free guard.
    "task_template_shared_v15.txt": "5b607903172352e82f8fbbac986d96ef",
}
SCRUBBED_FORBIDDEN = ("== FULL RUBRIC", "Good:", "Bad:", "rubric", "graded", "grading", "Category weights",
                      "professional-services firm")

# Re-cuts that carry rubric text BY DESIGN, so they cannot sit in SCRUBBED_MD5:
# whole-file md5 pins with their own message. v14 (ablation arm "rubric added
# back, no house standards") = v9 byte-identical; v9's rubric-section guard in
# prompt_file_paths applies to it as well. Never edit by hand.
RECUT_MD5 = {
    "task_template_shared_v14.txt": "7610bdae4c5eba21984ebcc9b19e3886",  # == md5 of task_template_shared_v9.txt
}

# Extra files a template asks the agent to read, staged into the WORKSPACE
# ROOT (not starting_files/) as (source name under prompts/, name in the
# workspace). They are prompt material: snapshotted with the prompts, listed
# under WORKSPACE FILES, hashed into the manifest, and recorded in
# extra_configs.prompt_extras.
TEMPLATE_EXTRAS = {
    "v11": [("house_standards_v1.md", "HOUSE_STANDARDS.md")],
    "v13": [("house_standards_v1.md", "HOUSE_STANDARDS.md")],
}


def _md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def template_name(version: str) -> str:
    return f"task_template_shared_{version}.txt"


def parse_prompt_version(system_name: str, template: str) -> int:
    sys_v = int(re.search(r"_v(\d+)\.txt$", system_name).group(1))
    tpl_v = int(re.search(r"_v(\d+)\.txt$", template).group(1))
    return sys_v * 100 + tpl_v


def prompt_file_paths(cfg: RunConfig) -> tuple[Path, Path]:
    system_path = PROMPTS_DIR / cfg.system_prompt
    tpl_path = PROMPTS_DIR / template_name(cfg.template_version)
    for p in (system_path, tpl_path):
        if not p.exists():
            raise FileNotFoundError(p)
    if tpl_path.name == "task_template_shared_v8.txt":
        text = tpl_path.read_text()
        idx = text.index(V8_RUBRIC_MARKER)
        rubric = text[idx:idx + V8_RUBRIC_LEN]
        if hashlib.md5(rubric.encode()).hexdigest() != V8_RUBRIC_MD5:
            raise RuntimeError(
                "v8 rubric section no longer matches the GUI original — do not edit it"
            )
    if tpl_path.name in RECUT_MD5 and _md5(tpl_path) != RECUT_MD5[tpl_path.name]:
        raise RuntimeError(
            f"{tpl_path.name} no longer matches its pinned md5 (a byte-identical re-cut) — never edit by hand"
        )
    if tpl_path.name in ("task_template_shared_v9.txt", "task_template_shared_v14.txt"):  # v14 = v9 text
        text = tpl_path.read_text()
        idx = text.index(V8_RUBRIC_MARKER)
        rubric = text[idx:idx + V9_RUBRIC_LEN]
        if hashlib.md5(rubric.encode()).hexdigest() != V9_RUBRIC_MD5:
            raise RuntimeError(
                f"{tpl_path.name}: v9 rubric section no longer matches the v3 GUI original — do not edit it"
            )
    if tpl_path.name == "task_template_shared_v12.txt":
        text = tpl_path.read_text()
        idx = text.index(V8_RUBRIC_MARKER)
        rubric = text[idx:idx + V12_RUBRIC_LEN]
        if hashlib.md5(rubric.encode()).hexdigest() != V12_RUBRIC_MD5:
            raise RuntimeError(
                "v12 rubric section no longer matches the v4 GUI original — do not edit it"
            )
    if tpl_path.name in SCRUBBED_MD5:
        text = tpl_path.read_text()
        hit = [w for w in SCRUBBED_FORBIDDEN if w.lower() in text.lower()]
        if hit:
            raise RuntimeError(f"{tpl_path.name} carries rubric wording {hit} — it must stay rubric-free")
        if _md5(tpl_path) != SCRUBBED_MD5[tpl_path.name]:
            raise RuntimeError(
                f"{tpl_path.name} no longer matches its pinned md5 — never edit by hand"
            )
    return system_path, tpl_path


def prompt_extra_paths(cfg: RunConfig) -> list[tuple[Path, str]]:
    """(source path, workspace name) for every extra file the template needs."""
    out = []
    for src_name, ws_name in TEMPLATE_EXTRAS.get(cfg.template_version, []):
        src = PROMPTS_DIR / src_name
        if not src.exists():
            raise FileNotFoundError(src)
        out.append((src, ws_name))
    return out


def prompt_extras_provenance(cfg: RunConfig) -> dict:
    """extra_configs.prompt_extras: {workspace name: {source, sha256}} (empty if none)."""
    return {
        ws_name: {"source": src.name, "sha256": hashlib.sha256(src.read_bytes()).hexdigest()}
        for src, ws_name in prompt_extra_paths(cfg)
    }


def build_prompt(cfg: RunConfig, spec: TaskSpec, workspace: Path, attempt=None) -> tuple[str, int]:
    """Write PROMPT.md (and any template extras) into the workspace; return
    (prompt_text, prompt_version). When `attempt` is given, extras are hashed
    into its manifest so validation knows they were seeded, not produced."""
    import shutil as _shutil
    system_path, tpl_path = prompt_file_paths(cfg)
    extras = prompt_extra_paths(cfg)
    for src, ws_name in extras:
        dest = workspace / ws_name
        _shutil.copy2(src, dest)
        if attempt is not None:
            from .workspace import sha256_file
            attempt.manifest[ws_name] = sha256_file(dest)
            (attempt.attempt_dir / "manifest.json").write_text(json.dumps(attempt.manifest, indent=2))
    listing = "\n".join(
        f"- starting_files/{p.name} ({p.stat().st_size:,} bytes)"
        for p in sorted((workspace / "starting_files").iterdir())
    )
    for _src, ws_name in extras:
        listing += f"\n- {ws_name} ({(workspace / ws_name).stat().st_size:,} bytes)"
    prompt = (
        f"{system_path.read_text()}\n"
        f"{tpl_path.read_text()}\n"
        f"WORKSPACE FILES:\n{listing}\n"
    )
    (workspace / "PROMPT.md").write_text(prompt)
    return prompt, parse_prompt_version(system_path.name, tpl_path.name)
