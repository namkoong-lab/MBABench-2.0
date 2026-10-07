import json
import re
from pathlib import Path

import yaml

from .excel_utils import NOGRADE_PREFIX, encode_file_to_base64


def check_letter(idx: int) -> str:
    """0-based check index -> Excel-style letter label (A..Z, AA, AB, ...).

    Rubric categories may exceed 26 checks (rubric_9's Formatting has 36),
    so labels continue past Z the way spreadsheet columns do. Every site
    that renders, derives, or maps check letters must use this function so
    prompt text, valid-letter lists, and name mappings agree.
    """
    label = ""
    idx += 1  # 1-based for the mod-26 arithmetic
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        label = chr(ord("A") + rem) + label
    return label


def load_template(template_path: str) -> list[dict]:
    """Load a YAML prompt template and return the judge_prompt message list."""
    with open(template_path) as f:
        data = yaml.safe_load(f)
    return data["judge_prompt"]


def render_content(
    content: str, params: dict, nullable_keys: set[str] | None = None
) -> str:
    """Replace {{ var }} placeholders with values from params dict.

    Uses regex instead of str.format() because template content contains
    literal { } in JSON examples that would break Python's formatter.

    Keys in nullable_keys that are missing or None render as empty string.
    """
    if nullable_keys is None:
        nullable_keys = set()

    def replacer(match):
        key = match.group(1).strip()
        if key in nullable_keys and (key not in params or params[key] is None):
            return ""
        if key not in params:
            raise KeyError(f"Missing template parameter: {key}")
        return str(params[key])

    return re.sub(r"\{\{\s*(\w+)\s*\}\}", replacer, content)


def _get_nullable_keys(entry: dict) -> set[str]:
    """Extract parameter names marked as nullable from an entry."""
    nullable = set()
    for p in entry.get("parameters", []):
        if isinstance(p, dict) and p.get("nullable", False):
            nullable.add(p["name"])
    return nullable


def _has_optional_data(entry: dict, kwargs: dict) -> bool:
    """Check whether an optional entry's required data is available in kwargs.

    Parameters can be dicts with {"name": str, "nullable": bool} or plain strings.
    An entry has its data if all non-nullable parameters are present and non-None.
    """
    params = entry.get("parameters", [])
    if not params:
        return False
    for p in params:
        if isinstance(p, dict):
            name = p["name"]
            nullable = p.get("nullable", False)
        else:
            name = p
            nullable = False
        if not nullable and (name not in kwargs or kwargs[name] is None):
            return False
    return True


def render_rubric_checks_list(
    items: list[dict], category: str = None, guidance: dict = None
) -> str:
    """Render an explicit check list into prompt-ready text.

    Letters are assigned by position in `items`, so a suitability-filtered
    list stays contiguous (A..N over the applicable checks) and matches the
    letter->name mapping built from the same list.

    `guidance` (utils.rubric_guidance.load_guidance) appends a "Guidance:"
    block under each check that has a note, and a category note above the
    checks; both need `category`. Omitted (the v3 / template_5 path), the
    output is byte-identical to before guidance existed.
    """
    blocks = []
    if guidance and category and category in guidance.get("categories", {}):
        blocks.append(
            f"Category guidance for {category}:\n"
            f"{guidance['categories'][category]}"
        )
    for i, item in enumerate(items):
        letter = check_letter(i)
        text = (
            f"Check {letter}: {item['description']}\n"
            f"Pass: {item['good']}\n"
            f"Fail: {item['bad']}"
        )
        if guidance and category:
            note = guidance.get("checks", {}).get((category, item["name"]))
            if note:
                text += f"\nGuidance: {note}"
        blocks.append(text)
    return "\n\n".join(blocks)


def numbered_rubric_checks(rubric: dict) -> list[tuple[int, str, dict]]:
    """Flatten a rubric into (global_no, category, check) triples.

    The numbering is the rubric's flattened order — category key order, then
    within-category order — i.e. exactly the 1..N sequence the suitability
    annotations' `no` field is validated against
    (utils.rubric_suitability.validate_annotation). IDs are therefore stable
    across tasks; suitability gating removes entries WITHOUT renumbering, so
    a gated list has gaps by design.
    """
    return [
        (no, cat, check)
        for no, (cat, check) in enumerate(
            ((cat, c) for cat, checks in rubric.items() for c in checks), start=1
        )
    ]


def render_rubric_checks_flat(
    numbered_items: list[tuple[int, str, dict]],
    guidance: dict = None,
    guidance_style: str = "footnote",
    category_extras: dict = None,
) -> str:
    """Render globally-numbered checks for the single-pass judge.

    Unlike the per-category renderer, the surrounding prompt supplies no
    category frame — so every check carries its number, [category] and name
    inline (`Check 17 [Formatting] Blue font for hardcoded inputs: ...`).
    Category guidance notes are rendered above the first check of their
    category; per-check guidance goes under its check.

    `guidance_style` (judge v6, item C): "footnote" renders per-check notes
    as an advisory "Guidance:" line (template_7, frozen); "standard" folds
    them into the check's standard as a "Standard:" line and labels the
    category note "Category standard" — the wording the template_8 system
    prompt declares part of the pass/fail definition. `category_extras`
    ({category: text}) appends extra binding text under a category note
    (template_8 puts the answer-equivalence rulebook under Accuracy).
    """
    blocks = []
    seen_categories = set()
    hardened = guidance_style == "standard"
    for no, category, item in numbered_items:
        if category not in seen_categories:
            note = (guidance or {}).get("categories", {}).get(category)
            extra = (category_extras or {}).get(category)
            if note or extra:
                head = (
                    f"Category standard for {category} (binding for every check below):"
                    if hardened
                    else f"Category guidance for {category}:"
                )
                body = "\n".join(t for t in (note, extra) if t)
                blocks.append(f"{head}\n{body}")
        seen_categories.add(category)
        text = (
            f"Check {no} [{category}] {item['name']}: {item['description']}\n"
            f"Pass: {item['good']}\n"
            f"Fail: {item['bad']}"
        )
        if guidance:
            note = guidance.get("checks", {}).get((category, item["name"]))
            if note:
                text += f"\n{'Standard' if hardened else 'Guidance'}: {note}"
        blocks.append(text)
    return "\n\n".join(blocks)


def render_rubric_checks(rubric_path: str, category: str) -> str:
    """Render a rubric category's checks into prompt-ready text.

    Each check becomes:
        Check A: {description}
        Pass: {good}
        Fail: {bad}

    Checks are separated by a blank line.
    """
    with open(rubric_path) as f:
        rubric = json.load(f)
    return render_rubric_checks_list(rubric[category])


def compile_prompt(template_path: str, **kwargs) -> list[list[dict]]:
    """Compile a YAML prompt template into staged message lists.

    Walks the template entries, renders msg content and injects placeholder
    messages. Splits at each 'response' entry to create stages.

    Returns a list of stages. Each stage is a list of message dicts
    ({"role": str, "content": str}) that should be appended to the
    conversation before making an LLM call.

    len(return_value) == number of LLM calls needed.

    Caller usage::

        stages = compile_prompt(template_path, **kwargs)
        all_messages = []
        responses = []
        for stage_messages in stages:
            all_messages.extend(stage_messages)
            response = my_llm_call(all_messages)
            all_messages.append({"role": "assistant", "content": response})
            responses.append(response)
    """
    template = load_template(template_path)
    stages: list[list[dict]] = []
    current_stage: list[dict] = []

    for entry in template:
        entry_type = entry.get("type")

        if entry_type == "msg":
            if entry.get("optional") and not _has_optional_data(entry, kwargs):
                continue
            nullable_keys = _get_nullable_keys(entry)
            content = render_content(entry["content"], kwargs, nullable_keys)
            current_stage.append({"role": entry["role"], "content": content})

        elif entry_type == "placeholder":
            name = entry["name"]
            if entry.get("optional") and name not in kwargs:
                continue
            current_stage.extend(kwargs[name])

        elif entry_type == "prior_response":
            # Slot for a previous stage's assistant response, filled at runtime.
            current_stage.append({
                "role": "assistant",
                "content": None,
                "_prior_stage": entry["stage"],
            })

        elif entry_type == "response":
            stages.append(current_stage)
            current_stage = []

    if current_stage:
        stages.append(current_stage)

    return stages


### Message Builder Functions ################################################


def add_file_confirmation(messages, header, prompt=None):
    """Wrap file messages with confirmation request/response scaffolding."""
    confirmation_msg = {
        "role": "user",
        "content": "Do not evaluate the files yet. Just confirm that you have received them.",
    }
    messages = [confirmation_msg] + messages
    confirmation_request = {
        "role": "user",
        "content": f'Respond with {{"{header}": "confirmed"}} if files are received',
    }
    messages.append(confirmation_request)

    if prompt is not None:
        prompt = (
            "Do not evaluate the files yet. Just confirm that you have received them.\n"
            + prompt
            + f'Respond with {{"{header}": "confirmed"}} if files are received\n'
        )
    return messages, prompt


def format_file_section(header, files_dict, add_confirmation=False):
    """Format a section of files for the prompt and messages.

    Returns:
        tuple: (messages, prompt, file_sizes) where file_sizes is a dict mapping
               filename to character count for that file's content.
    """
    messages = []
    prompt = f"{header}\n"
    sheet_num = 1
    file_sizes = {}

    for name, file_path in files_dict.items():
        file_path = Path(file_path)
        suffix = file_path.suffix.lower()

        if name.endswith("_full.csv"):
            base_name = name.replace("_full.csv", "")

            # Detect nograde sheets (context-only, not to be graded)
            is_nograde = base_name.startswith(NOGRADE_PREFIX)
            if is_nograde:
                display_name = base_name[len(NOGRADE_PREFIX):]
            else:
                display_name = base_name

            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    file_content = f.read()
            except UnicodeDecodeError:
                base64_content, mime_type = encode_file_to_base64(file_path)
                file_content = (
                    f"[File content encoded as base64: {base64_content[:100]}...]"
                )

            if is_nograde:
                nograde_label = (
                    " [CONTEXT ONLY - DO NOT GRADE: This sheet is provided for "
                    "reference context only. Do NOT evaluate or grade this sheet.]"
                )
                text_content = (
                    f"Sheet {sheet_num}: {display_name} (data - CSV format)"
                    f"{nograde_label}:\n{file_content}"
                )
            else:
                text_content = (
                    f"Sheet {sheet_num}: {display_name} (data - CSV format):\n{file_content}"
                )
            messages.append(
                {"role": "user", "content": [{"type": "text", "text": text_content}]}
            )
            prompt += f"Sheet {sheet_num}: {display_name} (data): {name}\n"
            file_sizes[name] = len(text_content)
            sheet_num += 1

        elif name.endswith("_additional_format.txt"):
            base_name = name.replace("_additional_format.txt", "")

            # Detect nograde sheets for formatting files too
            is_nograde = base_name.startswith(NOGRADE_PREFIX)
            if is_nograde:
                display_name = base_name[len(NOGRADE_PREFIX):]
            else:
                display_name = base_name

            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    file_content = f.read()
            except UnicodeDecodeError:
                base64_content, mime_type = encode_file_to_base64(file_path)
                file_content = (
                    f"[File content encoded as base64: {base64_content[:100]}...]"
                )

            if is_nograde:
                nograde_label = " [CONTEXT ONLY - DO NOT GRADE]"
                text_content = f"    {display_name} (formatting){nograde_label}:\n{file_content}"
            else:
                text_content = f"    {display_name} (formatting):\n{file_content}"
            messages.append(
                {"role": "user", "content": [{"type": "text", "text": text_content}]}
            )
            prompt += f"    {display_name} (formatting): {name}\n"
            file_sizes[name] = len(text_content)
        else:
            if suffix in [".txt", ".csv"]:
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        file_content = f.read()
                    text_content = f"{name}:\n{file_content}"
                    messages.append(
                        {
                            "role": "user",
                            "content": [{"type": "text", "text": text_content}],
                        }
                    )
                    file_sizes[name] = len(text_content)
                except UnicodeDecodeError:
                    base64_content, mime_type = encode_file_to_base64(file_path)
                    messages.append(
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": f"{name}:"},
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:{mime_type};base64,{base64_content}"
                                    },
                                },
                            ],
                        }
                    )
                    file_sizes[name] = len(f"{name}:") + len(base64_content)
            else:
                base64_content, mime_type = encode_file_to_base64(file_path)
                messages.append(
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": f"{name}:"},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:{mime_type};base64,{base64_content}"
                                },
                            },
                        ],
                    }
                )
                file_sizes[name] = len(f"{name}:") + len(base64_content)
            prompt += f"{name}: {name}\n"

    if add_confirmation:
        messages, prompt = add_file_confirmation(messages, header, prompt)

    return messages, prompt, file_sizes


### Rubric Helper Functions #################################################


def build_check_name_mapping(rubric_json_path: str) -> dict:
    """Build a mapping from (category, check_letter) to rubric name.

    Args:
        rubric_json_path: Path to the rubric JSON file.

    Returns:
        dict: Mapping of (category, letter) -> name,
              e.g., {("Accuracy", "A"): "Final Calculations"}
    """
    mapping = {}
    try:
        with open(rubric_json_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Support both flat-list-with-category and grouped-by-category formats
        if isinstance(data, dict) and "rubric" in data:
            # Old format: {"rubric": [{"category": "Accuracy", "name": "...", ...}]}
            rubrics = data["rubric"]
            category_items: dict[str, list] = {}
            for rubric in rubrics:
                category = rubric.get("category")
                if category:
                    if category not in category_items:
                        category_items[category] = []
                    category_items[category].append(rubric)

            for category, items in category_items.items():
                for idx, item in enumerate(items):
                    if "name" in item:
                        mapping[(category, check_letter(idx))] = item["name"]
        else:
            # New format: {"Accuracy": [...], "Formula": [...], ...}
            for category, items in data.items():
                if isinstance(items, list):
                    for idx, item in enumerate(items):
                        if isinstance(item, dict) and "description" in item:
                            name = item.get("name", item["description"][:40])
                            mapping[(category, check_letter(idx))] = name
    except Exception as e:
        import logging

        logging.getLogger(__name__).warning(f"Could not build check name mapping: {e}")

    return mapping

