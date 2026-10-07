"""Cell write tools: set_cell_formula, edit_cells."""
import json
import re
from typing import Any, Dict, List

from ..core.libreoffice_bridge import _recalculate_with_libreoffice, _save_with_recalc
from ..core.shared_state import _lo_engine, mcp
from ..core.workbook_io import _get_file_path, _load_workbook, _load_workbook_view, _save_workbook_sync
from .. import formula_validator as _formula_validator
from ..helpers.cell_validation import _validate_cell_references
from ..helpers.formula_evaluation import _eval_formula
from ..helpers.type_inference import _infer_type


@mcp.tool()
async def edit_cells(filename: str, worksheet_name: str, cell_updates: List[Dict[str, Any]]) -> str:
    """Edit multiple cells with values and formulas.

    Args:
        filename: Name of the Excel file
        worksheet_name: Name of the worksheet
        cell_updates: List of dictionaries with 'cell' and 'value' keys
                     Example: [{"cell": "A1", "value": "Hello"}, {"cell": "B1", "value": 42}]

    Returns:
        Success message with updated cell information
    """
    try:
        wb = _load_workbook(filename)
        ws = wb[worksheet_name]

        updated_cells = []
        for update in cell_updates:
            cell_address = update["cell"]
            value = update["value"]

            # Block ANY string starting with = (Excel treats all as formulas).
            # For actual formulas: use set_cell_formula instead.
            # For text headers like "=== HEADER ===": remove the = prefix (use "HEADER" or "--- HEADER ---").
            if isinstance(value, str) and value.startswith('='):
                return json.dumps({
                    "success": False,
                    "error": f"Values starting with '=' are not allowed in edit_cells (Excel treats them as formulas). "
                             f"Rejected value: '{value}' in cell {cell_address}. "
                             f"For formulas: use set_cell_formula tool. "
                             f"For text/headers: remove the '=' prefix (e.g., use '--- HEADER ---' instead of '=== HEADER ===').",
                    "cell": cell_address,
                    "rejected_value": value,
                }, indent=2)

            cell = ws[cell_address]
            cell.value = value

            updated_cells.append({
                "cell": cell_address,
                "value": value,
                "type": "value"
            })

        # Save + recalculate + report via the shared helper — see
        # libreoffice_bridge._save_with_recalc for why every save must
        # recalculate (an openpyxl save drops every cached value).
        recalc_engine_info = _save_with_recalc(wb, filename)

        return f"Successfully updated {len(updated_cells)} cells in '{filename}:{worksheet_name}'\n" + \
               json.dumps({"updated_cells": updated_cells,
                           "recalc_engine": recalc_engine_info}, indent=2)
    except Exception as e:
        return f"Error editing cells: {str(e)}"


@mcp.tool()
async def set_cell_formula(filename: str, worksheet_name: str, cell: str, formula: str) -> str:
    """Set a cell formula with validation and immediate feedback.

    This tool provides enhanced formula setting with:
    - Automatic '=' prefix addition if missing
    - Formula syntax validation
    - Worksheet reference validation
    - Immediate calculated value (if available)
    - Clear error messages for debugging

    Args:
        filename: Name of the Excel file
        worksheet_name: Name of the worksheet
        cell: Cell address (e.g., "B4", "C10")
        formula: Formula to set (with or without '=' prefix)
                Examples: "B2*B3", "=B2*B3", "SUM(A1:A10)", "='Financial Model'!B10"

    Returns:
        JSON string with success status, formula, calculated value (if available), or error details
    """
    try:
        if not formula.startswith('='):
            formula = f"={formula}"

        formula_content = formula[1:] if formula.startswith('=') else formula

        placeholder_indicators = ['placeholder', 'Placeholder', 'PLACEHOLDER', 'TODO', 'FIXME', 'TEMP']
        if any(indicator in formula_content for indicator in placeholder_indicators):
            return json.dumps({
                "success": False,
                "cell": cell,
                "formula": formula,
                "error": f"Formula contains placeholder text. Please replace with actual calculation. Found placeholder indicators: {[p for p in placeholder_indicators if p in formula_content]}",
                "error_type": "PLACEHOLDER_ERROR"
            }, indent=2)

        formula_stripped = formula_content.strip()
        try:
            float(formula_stripped.replace(',', ''))
            return json.dumps({
                "success": False,
                "cell": cell,
                "formula": formula,
                "error": f"Formula is just a constant number ({formula_stripped}). Use edit_cells for constants, not set_cell_formula. Constants should not have '=' prefix.",
                "error_type": "CONSTANT_ERROR"
            }, indent=2)
        except ValueError:
            pass

        # "Just text" means the WHOLE formula is one quoted literal. The old test
        # (starts and ends with a quote) also refused real formulas such as
        # ="FY"&TEXT(B6,"yyyy")&"A" (Fable 5.1 canary, 2026-09-19).
        if re.fullmatch(r'"(?:[^"]|"")*"', formula_stripped) or re.fullmatch(r"'[^']*'", formula_stripped):
            return json.dumps({
                "success": False,
                "cell": cell,
                "formula": formula,
                "error": f"Formula is just text ({formula_stripped}). Use edit_cells for text, not set_cell_formula. Text should not have '=' prefix.",
                "error_type": "TEXT_ERROR"
            }, indent=2)

        wb = _load_workbook(filename)

        if worksheet_name not in wb.sheetnames:
            return json.dumps({
                "success": False,
                "cell": cell,
                "formula": formula,
                "error": f"Worksheet '{worksheet_name}' does not exist. Available worksheets: {', '.join(wb.sheetnames)}",
                "error_type": "WORKSHEET_NOT_FOUND"
            }, indent=2)

        ws = wb[worksheet_name]

        # Text literals are emptied first: ="ERROR!" is not a sheet named ERROR,
        # and ="Note (see "&A1 has no unbalanced bracket.
        formula_no_text = _formula_validator.strip_string_literals(formula)
        sheet_ref_pattern = r"'([^']+)'!|([A-Za-z_][A-Za-z0-9_\s]*!)(?!')"
        matches = re.findall(sheet_ref_pattern, formula_no_text)

        for match in matches:
            referenced_sheet = match[0] if match[0] else match[1].rstrip('!')
            if referenced_sheet and referenced_sheet not in wb.sheetnames:
                return json.dumps({
                    "success": False,
                    "cell": cell,
                    "formula": formula,
                    "error": f"Invalid worksheet reference: '{referenced_sheet}' does not exist. Available worksheets: {', '.join(wb.sheetnames)}",
                    "error_type": "REF_ERROR"
                }, indent=2)

        try:
            from excel_mcp_server import formula_validator

            validation_result = formula_validator.validate_formula(
                formula,
                cell=cell,
                worksheet=worksheet_name
            )

            if not validation_result["valid"]:
                error_details = []
                if validation_result["errors"]:
                    error_details.extend(validation_result["errors"])
                if validation_result["warnings"]:
                    error_details.extend([f"Warning: {w}" for w in validation_result["warnings"]])

                return json.dumps({
                    "success": False,
                    "cell": cell,
                    "formula": formula,
                    "error": "Formula validation failed:\n" + "\n".join(error_details),
                    "error_type": "VALIDATION_ERROR",
                    "validation_details": {
                        "errors": validation_result["errors"],
                        "warnings": validation_result["warnings"],
                        "functions_used": validation_result["functions_used"]
                    }
                }, indent=2)

        except Exception as validation_error:
            print(f"Warning: Formula validation failed with error: {validation_error}")

        if formula_no_text.count('(') != formula_no_text.count(')'):
            return json.dumps({
                "success": False,
                "cell": cell,
                "formula": formula,
                "error": "Formula syntax error: Mismatched parentheses",
                "error_type": "SYNTAX_ERROR"
            }, indent=2)

        if formula.rstrip().endswith(('+', '-', '*', '/', '=', '>', '<', ',')):
            return json.dumps({
                "success": False,
                "cell": cell,
                "formula": formula,
                "error": "Formula syntax error: Formula ends with an operator",
                "error_type": "SYNTAX_ERROR"
            }, indent=2)

        cell_ref_errors = _validate_cell_references(formula, wb, current_worksheet=worksheet_name)
        if cell_ref_errors:
            return json.dumps({
                "success": False,
                "cell": cell,
                "formula": formula,
                "error": f"Cell reference validation failed: {', '.join(cell_ref_errors)}",
                "error_type": "CELL_TYPE_ERROR"
            }, indent=2)

        # Post-2007 functions (XLOOKUP, IFS, LET, STDEV.S, ...) must be stored
        # with their "_xlfn." prefix or they are "#NAME?" in Excel and LibreOffice.
        stored_formula = _formula_validator.add_future_function_prefixes(formula)
        target_cell = ws[cell]
        target_cell.value = stored_formula

        _save_workbook_sync(wb, _get_file_path(filename))

        calculated_value = None
        value_from_fallback_eval = False
        # Which engine stands behind calculated_value; always stamped on the
        # response so transcripts record fallback-engine values as such.
        recalc_engine_info = {"engine": "fallback"}

        from ..core.shared_state import _lo_engine
        lo_active = bool(_lo_engine and _lo_engine.is_running)
        if lo_active:
            lo_result = _recalculate_with_libreoffice(filename)
            if lo_result["success"]:
                recalc_engine_info = {
                    "engine": "libreoffice",
                    "duration_ms": lo_result.get("duration_ms", 0),
                }
                try:
                    wb_view = _load_workbook_view(filename)
                    ws_view = wb_view[worksheet_name]
                    calculated_value = ws_view[cell].value
                except Exception:
                    pass
            else:
                recalc_engine_info = {
                    "engine": "fallback",
                    "libreoffice_error": lo_result.get("error"),
                }

        if calculated_value is None:
            try:
                wb_view = _load_workbook_view(filename)
                ws_view = wb_view[worksheet_name]
                calculated_value = ws_view[cell].value
            except Exception:
                pass

        if calculated_value is None and not lo_active:
            # Deliberate no-LibreOffice mode only: the limited evaluator may
            # stand in, and the response labels it below. With LibreOffice
            # active, a missing cached value stays "Not available" — never a
            # possibly-divergent evaluator number.
            try:
                wb_raw = _load_workbook(filename)
                ws_raw = wb_raw[worksheet_name]
                calculated_value = _eval_formula(formula, wb_raw, ws_raw, {})
                value_from_fallback_eval = calculated_value is not None
            except Exception:
                pass

        response = {
            "success": True,
            "cell": cell,
            "formula": formula,
            "note": "Formula set successfully"
        }
        if stored_formula != formula:
            response["stored_as"] = stored_formula
            response["note"] += ("; newer functions are stored with the file format's _xlfn./_xlpm. "
                                 "prefixes (Excel shows them without) - expect them when reading formulas back")

        if calculated_value is not None:
            response["calculated_value"] = calculated_value
            response["value_type"] = _infer_type(calculated_value)
            if value_from_fallback_eval:
                response["calculated_value_engine"] = "fallback_eval"
        else:
            response["calculated_value"] = "Not available (open file in Excel to calculate)"

        response["recalc_engine"] = recalc_engine_info

        return json.dumps(response, indent=2, default=str)

    except KeyError as e:
        return json.dumps({
            "success": False,
            "cell": cell,
            "formula": formula,
            "error": f"Invalid cell reference: {str(e)}",
            "error_type": "CELL_ERROR"
        }, indent=2)
    except Exception as e:
        return json.dumps({
            "success": False,
            "cell": cell,
            "formula": formula,
            "error": f"Error setting formula: {str(e)}",
            "error_type": "GENERAL_ERROR"
        }, indent=2)
