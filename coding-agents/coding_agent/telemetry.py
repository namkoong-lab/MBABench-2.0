"""Telemetry: parse the agent's transcript into per-turn token usage + totals.

Best-effort by design — a parse failure must never fail an attempt (the raw
transcript is always preserved). Cost comes only from the CLI's own report
(claude's result event); we never maintain a price table. Codex does not
report cost, so cost stays None there and tokens carry the signal.

Output: telemetry.json
  {"turns": [{"n":1,"input_tokens":..,"output_tokens":..,"cache_read_tokens":..}, ...],
   "totals": {...}, "cost_usd": float|None, "num_turns": int|None, "parser": str}
"""
import json
from pathlib import Path


def _claude_parse(lines):
    turns, totals, cost, num_turns = [], {}, None, None
    n = 0
    for obj in lines:
        if obj.get("type") == "assistant":
            usage = (obj.get("message") or {}).get("usage") or {}
            if usage:
                n += 1
                turns.append({
                    "n": n,
                    "input_tokens": usage.get("input_tokens", 0),
                    "output_tokens": usage.get("output_tokens", 0),
                    "cache_read_tokens": usage.get("cache_read_input_tokens", 0),
                    "cache_creation_tokens": usage.get("cache_creation_input_tokens", 0),
                })
        elif obj.get("type") == "result":
            cost = obj.get("total_cost_usd")
            num_turns = obj.get("num_turns")
            totals = obj.get("usage") or {}
    return turns, totals, cost, num_turns


def _codex_parse(lines):
    """Codex --json event shapes vary by version; scan for token-usage dicts."""
    turns, last_cumulative = [], None

    def find_usage(node):
        if isinstance(node, dict):
            if "input_tokens" in node and "output_tokens" in node:
                return node
            for value in node.values():
                found = find_usage(value)
                if found:
                    return found
        elif isinstance(node, list):
            for value in node:
                found = find_usage(value)
                if found:
                    return found
        return None

    n = 0
    for obj in lines:
        usage = find_usage(obj)
        if not usage:
            continue
        snapshot = {
            "input_tokens": usage.get("input_tokens", 0),
            "output_tokens": usage.get("output_tokens", 0),
            "cache_read_tokens": usage.get("cached_input_tokens", 0),
            "cache_write_tokens": usage.get("cache_write_input_tokens", 0),
            "reasoning_output_tokens": usage.get("reasoning_output_tokens", 0),
        }
        if snapshot == last_cumulative:
            continue
        n += 1
        turns.append({"n": n, **snapshot})
        last_cumulative = snapshot
    # Codex reports cumulative counts; the last snapshot is the total.
    totals = dict(turns[-1]) if turns else {}
    totals.pop("n", None)
    return turns, totals, None, (len(turns) or None)


def parse_transcript(transcript_path: Path, cli: str) -> dict:
    result = {"turns": [], "totals": {}, "cost_usd": None, "num_turns": None,
              "parser": cli, "parse_error": None}
    try:
        lines = []
        with open(transcript_path, errors="replace") as f:
            for raw in f:
                raw = raw.strip()
                if not raw:
                    continue
                try:
                    lines.append(json.loads(raw))
                except json.JSONDecodeError:
                    continue
        parser = _claude_parse if cli == "claude" else _codex_parse
        turns, totals, cost, num_turns = parser(lines)
        result.update(turns=turns, totals=totals, cost_usd=cost, num_turns=num_turns)
    except Exception as e:  # noqa: BLE001 — telemetry must never sink an attempt
        result["parse_error"] = repr(e)
    return result


def write_telemetry(attempt_dir: Path, telemetry: dict) -> Path:
    path = attempt_dir / "telemetry.json"
    path.write_text(json.dumps(telemetry, indent=2))
    return path
