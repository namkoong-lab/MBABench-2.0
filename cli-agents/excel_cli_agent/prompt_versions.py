"""Shared prompt versioning config — single source of truth for the batch runner."""
import re
from pathlib import Path
from typing import List, Optional

PROMPTS_DIR = Path(__file__).resolve().parent / "prompts"

# Every registered prompt set: the system prompt and the (task-invariant)
# task template it pairs with, plus the files it ships into every workspace.
# All sets target the MBABench task set; v16 is the default and the version
# the recorded CLI cohorts ran under.
PROMPT_VERSIONS = {
    # v12: the v11 body with the 132-check/12-category MBABench rubric
    # (byte-exact from gui-agents prompts_v2/step2_build.txt) and one shared
    # 3-step task template. Generated, never hand-edited.
    "v12": {"system": "system_prompt_v12.txt", "template": "task_template_shared_v6.txt"},
    # v13: v12 + the Questions-sheet answer convention (gui prompts_v3/):
    # the starting workbook's 'Questions' sheet must be preserved in
    # solution.xlsx and answered with live formulas in column B (B2 down).
    # Rubric unchanged from v12. Recorded as prompt_version 1307.
    "v13": {"system": "system_prompt_v13.txt", "template": "task_template_shared_v7.txt"},
    # v14: v13 + the house financial-modelling standards (gui prompts_v4/):
    # House_Standards_v1.md is copied into every workspace and its full text
    # embedded in the model context; the prompts say to follow it, with the
    # case instructions and the rubric governing on conflict. Rubric
    # unchanged from v13. `attachments` are repo-root-relative paths the
    # runner ships with the task files — declared here, never in a batch
    # config, so the recorded prompt_version and the standards the agent saw
    # cannot disagree. Recorded as prompt_version 1408.
    "v14": {
        "system": "system_prompt_v14.txt",
        "template": "task_template_shared_v8.txt",
        "attachments": ["house_standards/House_Standards_v1.md"],
    },
    # v15: the rubric-scrubbed House-Standards set (2026-09-10) — v14 with
    # every rubric-derived passage removed (weights, conventions summary,
    # the 132-check block, the step-3 audit list, the "[rubric: ...]" tags
    # in the tool guidance). The agent gets the task, the ANSWERS mechanics,
    # the tool guidance and a pointer to HOUSE_STANDARDS.md — the name the
    # standards file is delivered under in the workspace
    # (`attachment_names`), matching the coding pipeline — with its full
    # text embedded in the context. Recorded as prompt_version 1509.
    "v15": {
        "system": "system_prompt_v15.txt",
        "template": "task_template_shared_v9.txt",
        "attachments": ["house_standards/House_Standards_v1.md"],
        "attachment_names": {"House_Standards_v1.md": "HOUSE_STANDARDS.md"},
    },
    # v16 (2026-09-19): v15 with the five passages of the tool manual that
    # contradicted the attached House Standards removed (outflows-positive
    # "Convention A", Calibri 11/12, 0.00% and no-decimal/no-dash number
    # codes, "font color, NOT cell fill") plus one sentence giving the
    # standards precedence over the manual's style advice. The standards'
    # values are not restated — the agent reads them from HOUSE_STANDARDS.md
    # as the GUI and coding agents do. Task template unchanged (v9). The
    # default; the recorded CLI cohorts carry prompt_version 1609.
    "v16": {
        "system": "system_prompt_v16.txt",
        "template": "task_template_shared_v9.txt",
        "attachments": ["house_standards/House_Standards_v1.md"],
        "attachment_names": {"House_Standards_v1.md": "HOUSE_STANDARDS.md"},
    },
}
# Model-specific SYSTEM PROMPT VARIANTS of a registered prompt set. Keyed by
# (prompt version, exact model id as the identity sends it). A variant is the
# set's system prompt with only its response contract changed, so the attempt
# still records the set's prompt_version (parse_prompt_version reads the
# _v16 of "system_prompt_v16_gemini-3.8-flash.txt" and yields 1609); the
# runner writes the file actually used into extra_configs.system_prompt_file.
# One entry (2026-09-22): Gemini 3.8 Flash through Forge answers with native
# function calls, so its v16 variant asks for those instead of the JSON
# actions batch - see models_config.GEMINI_TOOL_CALL_MODELS. Every other
# model gets the set's own system prompt, byte for byte.
MODEL_SYSTEM_PROMPT_VARIANTS = {
    ("v16", "tensorblock/gemini-3.8-flash"): "system_prompt_v16_gemini-3.8-flash.txt",
}

# What a batch config gets when it names no prompt_version.
DEFAULT_PROMPT_VERSION = "v16"


def attachments_for(prompt_ver: str) -> List[str]:
    """Repo-root-relative paths of the files this prompt set ships with
    every workspace (see repo_config.resolve_attachments for the absolute
    paths). [] for versions that attach nothing."""
    return list(PROMPT_VERSIONS[prompt_ver].get("attachments", []))


def attachment_names_for(prompt_ver: str) -> dict:
    """{source basename: name delivered into the workspace} for this prompt
    set. A source not listed keeps its own name. Declared on the version so
    the prompt's directive ("HOUSE_STANDARDS.md") and the file the agent
    finds cannot disagree."""
    return dict(PROMPT_VERSIONS[prompt_ver].get("attachment_names", {}))


def system_prompt_file(prompt_ver: str, model: Optional[str] = None, require_variant: bool = False) -> str:
    """The system prompt file this model runs the prompt set with: the
    set's own file, or its registered variant for exactly this model id
    (MODEL_SYSTEM_PROMPT_VARIANTS). Every model without a variant, and every
    call without a model, gets the set's file.

    require_variant (the runner passes models_config.uses_gemini_tool_calls
    (model)): a native-function-call model on a prompt set with no variant
    for it would be handed the JSON contract its tools contradict, so that
    refuses to run instead of silently running the set's own prompt."""
    variant = MODEL_SYSTEM_PROMPT_VARIANTS.get((prompt_ver, model or ""))
    if variant:
        return variant
    if require_variant:
        registered = sorted(v for v, m in MODEL_SYSTEM_PROMPT_VARIANTS if m == model)
        raise ValueError(
            f"{model} answers with native function calls and needs a system prompt variant, but "
            f"prompt_version {prompt_ver} has none for it (MODEL_SYSTEM_PROMPT_VARIANTS). "
            f"Versions with one: {registered or 'none'}.")
    return PROMPT_VERSIONS[prompt_ver]["system"]


def template_file(prompt_ver: str) -> str:
    """The task template file of a prompt set (one shared file per set)."""
    return PROMPT_VERSIONS[prompt_ver]["template"]


def parse_prompt_version(system_path: Path, template_path: Path) -> int:
    """Compute prompt_version: system_v{X} * 100 + template_v{Y}. A model
    variant of a system prompt ("system_prompt_v16_gemini-3.8-flash.txt")
    carries its set's version: the _v16 before the variant suffix."""
    sys_match = re.search(r'_v(\d+)(?:_[^/]*)?\.txt$', system_path.name)
    tpl_match = re.search(r'_v(\d+)\.txt$', template_path.name)
    sys_ver = int(sys_match.group(1)) if sys_match else 0
    tpl_ver = int(tpl_match.group(1)) if tpl_match else 0
    return sys_ver * 100 + tpl_ver
