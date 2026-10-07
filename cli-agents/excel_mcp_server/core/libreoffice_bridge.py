"""LibreOffice recalculation bridge."""
from .workbook_io import _get_file_path, _save_workbook_sync


def _recalculate_with_libreoffice(filename: str) -> dict:
    """Trigger LibreOffice recalculation for the given file.

    Returns dict with keys: success, duration_ms, error
    """
    from .shared_state import _lo_engine
    if _lo_engine is None or not _lo_engine.is_running:
        return {"success": False, "duration_ms": 0, "error": "LibreOffice engine not available"}

    file_path = str(_get_file_path(filename))
    return _lo_engine.recalculate(file_path)


def _save_with_recalc(wb, filename: str) -> dict:
    """Save a workbook, refresh its cached values, and report the outcome.

    EVERY openpyxl save drops EVERY formula cell's cached value, so any
    tool that saves must recalculate afterwards (or say loudly that it
    could not). Before this helper, tools that saved without recalcing
    (freeze_panes, format_cells, create/delete_worksheet, copy_file) left
    the whole file uncached whenever one of them was the run's last write —
    and the judge reads cached values (data_only=True), so a final cosmetic
    touch could silently zero an otherwise complete model.

    Returns the same engine-info shape edit_cells/set_cell_formula stamp on
    their responses: {"engine": "libreoffice", "duration_ms": ...} on
    success, {"engine": "fallback", ...} (with libreoffice_error + warning
    on a failed recalc) otherwise.
    """
    _save_workbook_sync(wb, _get_file_path(filename))
    info = {"engine": "fallback"}
    from .shared_state import _lo_engine
    if _lo_engine and _lo_engine.is_running:
        lo = _recalculate_with_libreoffice(filename)
        if lo["success"]:
            info = {"engine": "libreoffice", "duration_ms": lo.get("duration_ms", 0)}
        else:
            info = {
                "engine": "fallback",
                "libreoffice_error": lo.get("error"),
                "warning": "Recalculation FAILED — cached values in this "
                           "file are stale or missing until a "
                           "recalculation succeeds",
            }
    return info
