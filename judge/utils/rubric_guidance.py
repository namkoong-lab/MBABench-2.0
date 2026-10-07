"""Judge-only grading guidance, layered onto rubric checks at prompt time.

Why this file exists (2026-09-01, maintainer): repeat gradings of the same
attempt flip on checks whose *scope* is ambiguous — does formatting count
against Accuracy? do hardcodes inherited from the starting workbook count
against Flexibility? The guidance notes answer exactly those scope
questions so the judge applies one consistent rule instead of coin-flipping.

Why it is a separate file and NOT part of rubric_9.json:
  - rubric_9.json was generated from the canonical checklist workbook
    together with the agent-facing prompt blocks (prompts_v3/step2_build.txt
    and its mirrors carry the same text), so the rubric agents see and the
    rubric the judge grades cannot drift apart. The JSON is frozen: tests
    pin its text against the prompts.
  - Putting the guidance in the agents' prompt instead would change what
    agents are told — new prompt versions, re-run waves. These notes change
    how the JUDGE reads the rubric, not what agents are asked to do.

The guidance file is a sibling of the rubric it annotates
(<rubric_stem>_guidance.yaml, e.g. rubric_9_guidance.yaml). Structure:

    general: |            # overall rules, rendered once per prompt
      ...
    categories:           # per-category notes, rendered above the checks
      Accuracy: |
        ...
    checks:               # per-check notes, rendered under the check text
      - category: Accuracy
        name: Final calculation accuracy
        note: |
          ...

Loading validates every (category, name) against the rubric JSON and fails
loudly on drift, so a rubric revision that renames a check cannot silently
orphan its guidance. A rubric with no guidance sibling loads as None and
everything renders exactly as before.
"""

import json
from pathlib import Path

import yaml

try:
    from .logger import logger
except ImportError:  # imported as a bare module (utils/ on sys.path)
    from logger import logger


class GuidanceError(Exception):
    """The guidance file does not match the rubric it annotates."""


def guidance_path_for(rubric_path) -> Path:
    """The guidance sibling for a rubric JSON (may not exist)."""
    p = Path(rubric_path)
    return p.with_name(f"{p.stem}_guidance.yaml")


def load_guidance(rubric_path) -> dict | None:
    """Load and validate the guidance sibling of *rubric_path*.

    Returns {"general": str|None,
             "categories": {category: str},
             "checks": {(category, name): str}}
    or None when the rubric has no guidance file.

    Raises GuidanceError when a note names a category or check that the
    rubric does not contain — guidance must never silently detach from the
    rubric text it annotates.
    """
    gpath = guidance_path_for(rubric_path)
    if not gpath.exists():
        return None

    with open(rubric_path, "r", encoding="utf-8") as f:
        rubric = json.load(f)
    valid_pairs = {
        (cat, c["name"]) for cat, checks in rubric.items() for c in checks
    }
    valid_categories = set(rubric.keys())

    with open(gpath, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    general = raw.get("general")
    if general is not None:
        general = str(general).strip()

    categories: dict[str, str] = {}
    for cat, note in (raw.get("categories") or {}).items():
        if cat not in valid_categories:
            raise GuidanceError(
                f"{gpath.name}: category note for unknown category {cat!r}. "
                f"Rubric categories: {sorted(valid_categories)}"
            )
        categories[cat] = str(note).strip()

    checks: dict[tuple[str, str], str] = {}
    for entry in raw.get("checks") or []:
        cat = entry.get("category")
        name = entry.get("name")
        note = entry.get("note")
        if not cat or not name or not note:
            raise GuidanceError(
                f"{gpath.name}: every checks entry needs category, name and "
                f"note (got {entry!r})"
            )
        if (cat, name) not in valid_pairs:
            raise GuidanceError(
                f"{gpath.name}: note for unknown check ({cat!r}, {name!r}). "
                f"The rubric was probably revised — update the guidance file "
                f"to match before grading."
            )
        if (cat, name) in checks:
            raise GuidanceError(
                f"{gpath.name}: duplicate note for ({cat!r}, {name!r})"
            )
        checks[(cat, name)] = str(note).strip()

    logger.info(
        f"  Grading guidance loaded from {gpath.name}: "
        f"{len(checks)} check note(s), {len(categories)} category note(s), "
        f"general={'yes' if general else 'no'}"
    )
    return {"general": general, "categories": categories, "checks": checks}
