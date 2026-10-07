"""
MCP Client for Excel Operations (Synchronous Implementation)
Provides interface to Excel MCP Server tools with proper subprocess timeout handling.

This implementation uses direct subprocess communication instead of async MCP library,
which allows proper timeout enforcement that actually kills stuck processes.
"""
import collections
import json
import subprocess
import sys
import threading
import time
from typing import Any, Dict, List, Optional
from pathlib import Path


class ExcelMCPClient:
    """Synchronous MCP client with proper subprocess timeout handling."""

    # Per-tool timeout configuration (seconds)
    TOOL_TIMEOUTS = {
        "set_cell_formula": 90,         # Medium-heavy: includes LibreOffice recalculation
        "edit_cells": 90,               # Medium-heavy: includes LibreOffice recalculation
        "create_file": 30,              # Light: file creation
        "get_cell_range": 60,           # Medium: data read
        "scan_worksheet_structure": 60, # Medium: structure scan
        "search_worksheet": 60,         # Medium: search
        "summarize_workbook_context": 90,  # Medium-heavy: workbook summary
        "describe_worksheet": 90,       # Medium-heavy: LLM description
        "default": 120                  # Default for unlisted tools
    }

    CONNECT_TIMEOUT = 60  # Timeout for server startup

    def __init__(self, server_script_path: str, storage_path: str = "./excel_files",
                 server_args: Optional[List[str]] = None):
        self.server_script_path = Path(server_script_path)
        self.storage_path = storage_path
        # Extra argv appended to the server command (e.g. --no-libreoffice,
        # --allow-recalc-fallback).
        self.server_args: List[str] = list(server_args or [])
        self.process: Optional[subprocess.Popen] = None
        self.available_tools: List[str] = []
        # The server's own tool list (name, description, inputSchema) - declared
        # to Gemini 3.8 Flash as functions (see ExcelTaskExecutor._gemini_tool_declarations).
        self.tool_schemas: List[Dict[str, Any]] = []
        self.created_files: List[str] = []
        self._request_id = 0
        self._connected = False
        # Last lines the server wrote to stderr (see _drain_stderr).
        self._stderr_tail: collections.deque = collections.deque(maxlen=200)

    def _next_request_id(self) -> int:
        """Generate next JSON-RPC request ID."""
        self._request_id += 1
        return self._request_id

    def _send_request(self, method: str, params: Optional[Dict] = None, timeout: float = 60) -> Dict[str, Any]:
        """Send JSON-RPC request and wait for response with timeout.

        If timeout occurs, the subprocess is killed to prevent zombie processes.
        """
        if not self.process or self.process.poll() is not None:
            raise Exception("MCP server not running")

        request_id = self._next_request_id()
        request = {
            "jsonrpc": "2.0",
            "id": request_id,
            "method": method,
        }
        if params:
            request["params"] = params

        request_line = json.dumps(request) + "\n"

        try:
            # Write request to stdin
            self.process.stdin.write(request_line.encode())
            self.process.stdin.flush()

            # Read response with timeout using a thread
            response_line = self._read_line_with_timeout(timeout)

            if response_line is None:
                # Timeout occurred - kill the process
                tail = list(self._stderr_tail)[-10:]
                if tail:
                    print("   MCP server stderr (last lines):\n      " + "\n      ".join(tail))
                self._kill_process()
                raise TimeoutError(f"MCP request '{method}' timed out after {timeout}s")

            # Parse response
            response = json.loads(response_line)

            if "error" in response:
                error = response["error"]
                raise Exception(f"MCP error: {error.get('message', str(error))}")

            return response.get("result", {})

        except TimeoutError:
            raise
        except Exception as e:
            raise Exception(f"MCP request failed: {str(e)}")

    def _read_line_with_timeout(self, timeout: float) -> Optional[str]:
        """Read a line from stdout with timeout.

        Returns None if timeout occurs.
        """
        result = [None]
        exception = [None]

        def read_thread():
            try:
                line = self.process.stdout.readline()
                if line:
                    result[0] = line.decode().strip()
            except Exception as e:
                exception[0] = e

        thread = threading.Thread(target=read_thread)
        thread.daemon = True
        thread.start()
        thread.join(timeout=timeout)

        if thread.is_alive():
            # Timeout - thread is still blocking on readline
            return None

        if exception[0]:
            raise exception[0]

        return result[0]

    def _drain_stderr(self, stream) -> None:
        """Keep the server's stderr pipe empty, remembering its last lines.

        The server logs a ~43-byte line per request to stderr. The pipe was
        opened and never read, so its 64 KB buffer filled at the 1,505th
        call of a session and the server blocked on its next log write:
        the tool call hung until its 60-120 s timeout, the server was
        killed and the iteration lost (29 such timeouts in the v2 logs,
        every one at call 1,505/1,506; reproduced 2026-09-19).
        """
        rest = b""
        try:
            while True:
                chunk = stream.read(65536)  # raw pipe: returns what is there
                if not chunk:
                    break
                *lines, rest = (rest + chunk).split(b"\n")
                self._stderr_tail.extend(l.decode(errors="replace").rstrip() for l in lines)
        except (OSError, ValueError):
            pass  # pipe closed under us (server killed)
        if rest:
            self._stderr_tail.append(rest.decode(errors="replace").rstrip())

    def _kill_process(self):
        """Kill the MCP server subprocess."""
        if self.process:
            try:
                self.process.kill()
                self.process.wait(timeout=5)
            except Exception:
                pass
            finally:
                self.process = None
                self._connected = False

    def connect(self):
        """Start MCP server subprocess and initialize connection."""
        if self._connected and self.process and self.process.poll() is None:
            return  # Already connected

        try:
            # Start the MCP server subprocess
            self.process = subprocess.Popen(
                [sys.executable, str(self.server_script_path), self.storage_path, *self.server_args],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                bufsize=0  # Unbuffered
            )
            self._stderr_tail.clear()
            threading.Thread(target=self._drain_stderr, args=(self.process.stderr,),
                             name="mcp-stderr-drain", daemon=True).start()

            # Initialize the MCP session
            init_result = self._send_request("initialize", {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "excel-cli-agent",
                    "version": "1.0.0"
                }
            }, timeout=self.CONNECT_TIMEOUT)

            # Send initialized notification (no response expected, but we need to handle it)
            # The MCP protocol requires this after initialize
            notif = {"jsonrpc": "2.0", "method": "notifications/initialized"}
            self.process.stdin.write((json.dumps(notif) + "\n").encode())
            self.process.stdin.flush()

            # List available tools
            tools_result = self._send_request("tools/list", {}, timeout=self.CONNECT_TIMEOUT)

            if "tools" in tools_result:
                self.available_tools = [tool["name"] for tool in tools_result["tools"]]
                self.tool_schemas = list(tools_result["tools"])

            self._connected = True
            print(f"✅ Connected to Excel MCP Server")
            print(f"📊 Available tools ({len(self.available_tools)}): {', '.join(self.available_tools)}")

        except Exception as e:
            self._kill_process()
            raise Exception(f"Failed to connect to Excel MCP Server: {str(e)}")

    def disconnect(self):
        """Gracefully disconnect from MCP server."""
        if self.process:
            try:
                # Try graceful shutdown first
                self.process.stdin.close()
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self._kill_process()
            except Exception:
                self._kill_process()
            finally:
                self.process = None
                self._connected = False
        print("🔌 Disconnected from Excel MCP Server")

    def call_tool(self, tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Call an MCP tool with proper timeout handling."""
        if not self._connected:
            self.connect()

        if tool_name not in self.available_tools:
            raise ValueError(f"Tool '{tool_name}' not available. Available tools: {self.available_tools}")

        # Log the tool call
        args_str = json.dumps(arguments, indent=2)
        print(f"🔧 Calling tool: {tool_name}")
        print(f"   Arguments: {args_str}")

        timeout = self.TOOL_TIMEOUTS.get(tool_name, self.TOOL_TIMEOUTS["default"])

        try:
            result = self._send_request("tools/call", {
                "name": tool_name,
                "arguments": arguments
            }, timeout=timeout)

            # Process result - MCP returns content array
            if "content" in result and result["content"]:
                content = result["content"][0]
                text = content.get("text", str(content))

                # Try to parse as JSON
                try:
                    parsed_data = json.loads(text)
                    response = {
                        "success": True,
                        "tool": tool_name,
                        "arguments": arguments,
                        "result": parsed_data,
                        "raw_text": text
                    }
                except json.JSONDecodeError:
                    response = {
                        "success": True,
                        "tool": tool_name,
                        "arguments": arguments,
                        "result": text,
                        "raw_text": text
                    }

                # Track created files
                if tool_name == "create_file" and "filename" in arguments:
                    file_path = Path(self.storage_path) / arguments["filename"]
                    abs_path = str(file_path.absolute())
                    if abs_path not in self.created_files:
                        self.created_files.append(abs_path)

                return response
            else:
                return {
                    "success": True,
                    "tool": tool_name,
                    "arguments": arguments,
                    "result": result,
                    "raw_text": str(result)
                }

        except TimeoutError as e:
            print(f"⚠️ MCP tool '{tool_name}' timed out after {timeout}s - process killed")
            # Reconnect for next call
            self.connect()
            return {
                "success": False,
                "tool": tool_name,
                "arguments": arguments,
                "error": f"MCP tool call timed out after {timeout}s - operation took too long"
            }
        except Exception as e:
            return {
                "success": False,
                "tool": tool_name,
                "arguments": arguments,
                "error": str(e)
            }

    # Convenience methods - now synchronous
    def create_file(self, filename: str, worksheets: Optional[List[str]] = None) -> Dict[str, Any]:
        """Create a new Excel file"""
        return self.call_tool("create_file", {
            "filename": filename,
            "worksheets": worksheets or ["Sheet1"]
        })

    def list_files(self) -> Dict[str, Any]:
        """List all Excel files"""
        return self.call_tool("list_files", {})

    def get_file_metadata(self, filename: str) -> Dict[str, Any]:
        """Get file metadata"""
        return self.call_tool("get_file_metadata", {"filename": filename})

    def delete_file(self, filename: str) -> Dict[str, Any]:
        """Delete an Excel file"""
        return self.call_tool("delete_file", {"filename": filename})

    def list_worksheets(self, filename: str) -> Dict[str, Any]:
        """List worksheets in a file"""
        return self.call_tool("list_worksheets", {"filename": filename})

    def create_worksheet(self, filename: str, worksheet_name: str) -> Dict[str, Any]:
        """Create a new worksheet"""
        return self.call_tool("create_worksheet", {
            "filename": filename,
            "worksheet_name": worksheet_name
        })

    def delete_worksheet(self, filename: str, worksheet_name: str) -> Dict[str, Any]:
        """Delete a worksheet"""
        return self.call_tool("delete_worksheet", {
            "filename": filename,
            "worksheet_name": worksheet_name
        })

    def get_cell_range(self, filename: str, worksheet_name: str, range_address: str, mode: Optional[str] = None) -> Dict[str, Any]:
        """Get values from a cell range"""
        args: Dict[str, Any] = {
            "filename": filename,
            "worksheet_name": worksheet_name,
            "range_address": range_address,
        }
        if mode:
            args["mode"] = mode
        return self.call_tool("get_cell_range", args)

    def get_formula(self, filename: str, worksheet_name: str, cell_address: str) -> Dict[str, Any]:
        """Get formula from a cell"""
        return self.call_tool("get_formula", {
            "filename": filename,
            "worksheet_name": worksheet_name,
            "cell_address": cell_address
        })

    def edit_cells(self, filename: str, worksheet_name: str, cell_updates: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Edit multiple cells"""
        return self.call_tool("edit_cells", {
            "filename": filename,
            "worksheet_name": worksheet_name,
            "cell_updates": cell_updates
        })

    def get_used_range(self, filename: str, worksheet_name: str, max_rows: Optional[int] = None, max_cols: Optional[int] = None) -> Dict[str, Any]:
        """Get actual used range bounds of a worksheet."""
        args: Dict[str, Any] = {"filename": filename, "worksheet_name": worksheet_name}
        if max_rows is not None:
            args["max_rows"] = max_rows
        if max_cols is not None:
            args["max_cols"] = max_cols
        return self.call_tool("get_used_range", args)

    def scan_worksheet_structure(self, filename: str, worksheet_name: str, row_cap: int = 2000, col_cap: int = 100) -> Dict[str, Any]:
        """Scan worksheet and return a compact structural summary."""
        return self.call_tool("scan_worksheet_structure", {
            "filename": filename,
            "worksheet_name": worksheet_name,
            "row_cap": row_cap,
            "col_cap": col_cap,
        })

    def search_worksheet(self, filename: str, worksheet_name: str, query: str, match_type: str = "substring", case_sensitive: bool = False, max_results: int = 200) -> Dict[str, Any]:
        """Search values in a worksheet."""
        return self.call_tool("search_worksheet", {
            "filename": filename,
            "worksheet_name": worksheet_name,
            "query": query,
            "match_type": match_type,
            "case_sensitive": case_sensitive,
            "max_results": max_results,
        })

    def describe_worksheet(self, filename: str, worksheet_name: str, focus: Optional[str] = None, row_cap: int = 2000, col_cap: int = 100, model: Optional[str] = None) -> Dict[str, Any]:
        """LLM-powered description of worksheet with ranges."""
        args: Dict[str, Any] = {
            "filename": filename,
            "worksheet_name": worksheet_name,
            "row_cap": row_cap,
            "col_cap": col_cap,
        }
        if focus:
            args["focus"] = focus
        if model:
            args["model"] = model
        return self.call_tool("describe_worksheet", args)

    def summarize_workbook_context(self, filename: str, row_cap: int = 2000, col_cap: int = 120, max_cells: int = 150_000) -> Dict[str, Any]:
        """Summarize workbook context across all sheets."""
        return self.call_tool("summarize_workbook_context", {
            "filename": filename,
            "row_cap": row_cap,
            "col_cap": col_cap,
            "max_cells": max_cells,
        })

    def set_cell_value(self, filename: str, worksheet_name: str, cell_address: str, value: Any) -> Dict[str, Any]:
        """Set a single cell value"""
        return self.edit_cells(filename, worksheet_name, [
            {"cell": cell_address, "value": value}
        ])

    def set_cell_formula(self, filename: str, worksheet_name: str, cell_address: str, formula: str) -> Dict[str, Any]:
        """Set a single cell formula"""
        if not formula.startswith('='):
            formula = '=' + formula
        return self.edit_cells(filename, worksheet_name, [
            {"cell": cell_address, "value": formula}
        ])

    def get_cell_value(self, filename: str, worksheet_name: str, cell_address: str) -> Dict[str, Any]:
        """Get a single cell value"""
        res = self.get_cell_range(filename, worksheet_name, cell_address)
        try:
            if res.get('success') and isinstance(res.get('result'), dict):
                rv = res['result']
                vv = rv.get('view_values')
                if vv is not None:
                    rv['values'] = vv
        except Exception:
            pass
        return res

    def is_connected(self) -> bool:
        """Check if connected to MCP server"""
        return self._connected and self.process is not None and self.process.poll() is None

    def get_available_tools(self) -> List[str]:
        """Get list of available tools"""
        return self.available_tools.copy()

    def get_created_files(self) -> List[str]:
        """Get list of files created during this session"""
        return self.created_files.copy()

    def clear_created_files(self) -> None:
        """Clear the list of created files"""
        self.created_files = []

    # Context manager support (synchronous)
    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
