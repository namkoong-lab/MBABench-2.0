"""Shared state: MCP server instance, storage paths, and global engine references."""
import os
from pathlib import Path
from typing import Optional

from mcp.server.fastmcp import FastMCP
from openai import OpenAI

# Initialize the MCP server
mcp = FastMCP("excel-mcp-server")

# Global storage path - will be set by command line argument
STORAGE_PATH = Path("./excel_files")
ISSUES_PATH = Path(os.environ.get("EXCEL_AGENT_ISSUES_DIR", Path(__file__).resolve().parent.parent / "issues"))

# Optional OpenAI client for descriptive analysis; initialized lazily
_openai_client: Optional[OpenAI] = None

# Global LibreOffice engine instance (typed loosely to avoid import dependency)
_lo_engine = None


def _get_openai_client() -> Optional[OpenAI]:
    global _openai_client
    if _openai_client is not None:
        return _openai_client
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None
    try:
        _openai_client = OpenAI(api_key=api_key)
        return _openai_client
    except Exception:
        return None
