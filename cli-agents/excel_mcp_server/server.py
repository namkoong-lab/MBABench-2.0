#!/usr/bin/env python3
"""
Excel MCP Server
Provides Excel manipulation tools via MCP protocol using FastMCP.

Tools are registered in excel_mcp_server/tools/ via @mcp.tool() decorators.
Shared state (mcp instance, storage path) lives in excel_mcp_server/core/shared_state.py.
"""
import sys
import argparse
from pathlib import Path

# Ensure the project root is on sys.path so `excel_mcp_server` is importable
# when this file is executed directly (e.g., `python excel_mcp_server/server.py`).
_project_root = str(Path(__file__).resolve().parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

# Import shared state (mcp instance, STORAGE_PATH, etc.)
from excel_mcp_server.core import shared_state  # noqa: E402

# Import all tool modules to trigger @mcp.tool() registration
import excel_mcp_server.tools  # noqa: F401, E402

# LibreOffice recalculation engine: direct `soffice --convert-to` per recalc,
# no UNO. A plain package import — the engine has no import-time dependency on
# LibreOffice itself, so availability is decided at start(), loudly, not here.
from excel_mcp_server.libreoffice_calc import LibreOfficeCalcEngine  # noqa: E402


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="Excel MCP Server")
    parser.add_argument("storage_path", nargs="?", default="./excel_files",
                       help="Path to Excel files storage directory")
    parser.add_argument("--no-libreoffice", action="store_true",
                       help="Disable LibreOffice recalculation engine")
    parser.add_argument("--allow-recalc-fallback", action="store_true",
                       help="If the LibreOffice engine cannot start, keep running "
                            "with the limited _eval_formula fallback instead of "
                            "exiting. Without this flag an unavailable engine is "
                            "fatal, so a batch can never silently produce "
                            "fallback-engine attempts.")
    args = parser.parse_args()

    # Set storage path on shared state
    shared_state.STORAGE_PATH = Path(args.storage_path)
    shared_state.STORAGE_PATH.mkdir(exist_ok=True)

    print(f"Excel MCP Server starting with storage: {shared_state.STORAGE_PATH}", file=sys.stderr)

    if args.no_libreoffice:
        print("LibreOffice engine disabled by --no-libreoffice flag", file=sys.stderr)
    else:
        try:
            shared_state._lo_engine = LibreOfficeCalcEngine()
            shared_state._lo_engine.start()
            info = shared_state._lo_engine.info()
            print(f"LibreOffice recalc engine started: {info['soffice_path']} "
                  f"({info['soffice_version'] or 'version unknown'})", file=sys.stderr)
        except Exception as e:
            shared_state._lo_engine = None
            if args.allow_recalc_fallback:
                print(f"Warning: LibreOffice engine failed to start: {e}", file=sys.stderr)
                print("--allow-recalc-fallback set: continuing with the limited "
                      "_eval_formula fallback (single-cell evaluation, no cached "
                      "values saved for the judge).", file=sys.stderr)
            else:
                print(f"Error: LibreOffice engine failed to start: {e}", file=sys.stderr)
                print("Refusing to run with the degraded _eval_formula fallback. "
                      "Install LibreOffice (or set libreoffice_path in "
                      "<MBABench>/config/config.yaml / LIBREOFFICE_PATH), or pass "
                      "--allow-recalc-fallback / --no-libreoffice to opt out "
                      "explicitly.", file=sys.stderr)
                sys.exit(1)

    try:
        shared_state.mcp.run()
    finally:
        if shared_state._lo_engine:
            shared_state._lo_engine.stop()


if __name__ == "__main__":
    main()
