"""Worksheet management tools: list, create, delete."""
import json

from ..core.libreoffice_bridge import _save_with_recalc
from ..core.shared_state import mcp
from ..core.workbook_io import _load_workbook


@mcp.tool()
async def list_worksheets(filename: str) -> str:
    """List all worksheets in an Excel file.

    Args:
        filename: Name of the Excel file

    Returns:
        JSON string with worksheet information
    """
    try:
        wb = _load_workbook(filename)
        worksheets = []
        for ws in wb.worksheets:
            worksheets.append({
                "name": ws.title,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "is_active": ws == wb.active
            })

        return json.dumps(worksheets, indent=2)
    except Exception as e:
        return f"Error listing worksheets: {str(e)}"


@mcp.tool()
async def create_worksheet(filename: str, worksheet_name: str) -> str:
    """Create a new worksheet in an Excel file.

    Args:
        filename: Name of the Excel file
        worksheet_name: Name of the new worksheet

    Returns:
        Success or error message
    """
    try:
        wb = _load_workbook(filename)

        if worksheet_name in wb.sheetnames:
            return (
                f"ERROR: Worksheet '{worksheet_name}' already exists in '{filename}'!\n\n"
                f"create_worksheet is for NEW worksheets only.\n\n"
                f"WHY THIS MATTERS:\n"
                f"   - Calling create_worksheet on existing worksheet causes openpyxl to auto-rename\n"
                f"   - Example: create_worksheet('Q1') when Q1 exists -> openpyxl creates 'Q11'\n"
                f"   - Result: You end up with Q1, Q11, Q12, Q13... instead of just Q1\n\n"
                f"WHAT TO DO INSTEAD:\n"
                f"   - Use edit_cells('{filename}', '{worksheet_name}', ...) to modify existing worksheet\n"
                f"   - Use set_cell_formula('{filename}', '{worksheet_name}', ...) to add formulas\n"
                f"   - Use get_cell_range('{filename}', '{worksheet_name}', ...) to read data\n\n"
                f"TIP: Use list_worksheets('{filename}') to check existing worksheets first.\n"
                f"Current worksheets in '{filename}': {wb.sheetnames}"
            )

        wb.create_sheet(worksheet_name)
        recalc = _save_with_recalc(wb, filename)

        return (f"Worksheet '{worksheet_name}' created successfully in '{filename}'\n"
                f"recalc_engine: {json.dumps(recalc)}")
    except Exception as e:
        return f"Error creating worksheet: {str(e)}"


@mcp.tool()
async def delete_worksheet(filename: str, worksheet_name: str) -> str:
    """Delete a worksheet from an Excel file.

    Args:
        filename: Name of the Excel file
        worksheet_name: Name of the worksheet to delete

    Returns:
        Success or error message
    """
    try:
        wb = _load_workbook(filename)
        if worksheet_name in wb.sheetnames:
            wb.remove(wb[worksheet_name])
            recalc = _save_with_recalc(wb, filename)
            return (f"Worksheet '{worksheet_name}' deleted successfully from '{filename}'\n"
                    f"recalc_engine: {json.dumps(recalc)}")
        else:
            return f"Worksheet '{worksheet_name}' not found in '{filename}'"
    except Exception as e:
        return f"Error deleting worksheet: {str(e)}"
