"""Meta tools: report_mcp_issue, validate_formula, get_recalc_engine_info."""
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from ..core.shared_state import mcp, ISSUES_PATH
from ..core.workbook_io import _slugify


@mcp.tool()
async def report_mcp_issue(
    category: str,
    title: str,
    description: str,
    severity: str = "normal",
    task_id: Optional[str] = None,
    tool_name: Optional[str] = None,
    tool_args: Optional[Dict[str, Any]] = None,
    error: Optional[str] = None,
    context: Optional[Dict[str, Any]] = None,
) -> str:
    """Report an MCP server/interface issue by writing a file under issues/.

    Args:
        category: Short category like 'formula_support', 'tool_missing', 'bug'.
        title: Human-readable issue title.
        description: Detailed description of the problem.
        severity: One of 'low', 'normal', 'high', 'critical'.
        task_id: Optional task identifier from the agent.
        tool_name: Optional tool that triggered the issue.
        tool_args: Optional arguments used when the issue occurred.
        error: Optional error message observed.
        context: Optional extra diagnostic context.

    Returns:
        JSON string with {"success": true, "path": "issues/....md"} on success.
    """
    try:
        ISSUES_PATH.mkdir(parents=True, exist_ok=True)

        ts = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        slug = _slugify(title)
        category_slug = _slugify(category)
        filename = f"{ts}-{category_slug}-{slug}.md"
        file_path = ISSUES_PATH / filename

        payload: Dict[str, Any] = {
            "title": title,
            "category": category,
            "severity": severity,
            "task_id": task_id,
            "tool_name": tool_name,
            "tool_args": tool_args,
            "error": error,
            "description": description,
            "context": context or {},
            "timestamp_utc": ts,
        }

        md = []
        md.append(f"# {title}")
        md.append("")
        md.append(f"- Category: {category}")
        md.append(f"- Severity: {severity}")
        if task_id:
            md.append(f"- Task ID: {task_id}")
        if tool_name:
            md.append(f"- Tool: {tool_name}")
        md.append(f"- Timestamp (UTC): {ts}")
        md.append("")
        md.append("## Description")
        md.append(description.strip())
        if error:
            md.append("")
            md.append("## Error")
            md.append("```")
            md.append(str(error).strip())
            md.append("```")
        md.append("")
        md.append("## Context (JSON)")
        md.append("```json")
        md.append(json.dumps(payload, indent=2, default=str))
        md.append("```")
        md.append("")

        file_path.write_text("\n".join(md), encoding="utf-8")

        return json.dumps({
            "success": True,
            "message": "Issue recorded",
            "path": str(file_path),
            "filename": filename,
        })
    except Exception as e:
        return json.dumps({
            "success": False,
            "message": f"Failed to write issue: {str(e)}",
        })


@mcp.tool()
async def validate_formula(formula: str) -> str:
    """Validate an Excel formula for syntax and function name errors.

    Use this tool to check formulas BEFORE using set_cell_formula to catch:
    - Invalid Excel function names (e.g., SUMPMT should be CUMIPMT)
    - Potential undefined named ranges
    - Basic syntax errors (unbalanced parentheses, etc.)

    Args:
        formula: Excel formula string (should start with '=')

    Returns:
        JSON string with validation results:
        {
            "valid": bool,
            "errors": List[str],
            "warnings": List[str],
            "functions_used": List[str],
            "potential_names": List[str]
        }
    """
    try:
        from excel_mcp_server import formula_validator

        result = formula_validator.validate_formula(formula)

        return json.dumps(result, indent=2)

    except Exception as e:
        return json.dumps({
            "valid": False,
            "errors": [f"Validation error: {str(e)}"],
            "warnings": [],
            "functions_used": [],
            "potential_names": []
        }, indent=2)


@mcp.tool()
async def get_recalc_engine_info() -> str:
    """Report which formula recalculation engine this server is using.

    Harness-facing provenance tool: the batch runner records the answer on
    every attempt (extra_configs.recalc_engine) so fallback-engine runs can
    never be mistaken for LibreOffice runs.

    Returns:
        JSON string {"engine": "libreoffice", "soffice_path": ..., "soffice_source": ...,
        "soffice_version": ...} when LibreOffice recalculation is active, or
        {"engine": "fallback"} when running with the limited _eval_formula fallback.
    """
    from ..core.shared_state import _lo_engine

    if _lo_engine is not None and _lo_engine.is_running:
        return json.dumps(_lo_engine.info(), indent=2)
    return json.dumps({"engine": "fallback"}, indent=2)
