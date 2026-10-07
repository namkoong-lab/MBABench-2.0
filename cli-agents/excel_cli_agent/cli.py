#!/usr/bin/env python3
"""
Excel CLI Agent (Synchronous Implementation)
Interactive command-line interface for Excel automation
"""
import argparse
import sys
import os
from pathlib import Path
from typing import Optional, List, Dict, Any
import json
import readline  # For command history and editing
from dotenv import load_dotenv

from .mcp_client import ExcelMCPClient
from .task_executor import ExcelTaskExecutor, TaskStatus
from .batch_runner import dry_run_from_config, run_batch_from_config
from .models_config import DEFAULT_MAX_ITERATIONS, DEFAULT_MODEL
from .repo_config import repo_value


class ExcelCLIAgent:
    def __init__(self, server_path: str, storage_path: str, api_key: str, model: str = DEFAULT_MODEL, fresh_context_mode: bool = False, enhanced_excel_context: bool = True, enable_langfuse: bool = False, base_url: str = None, allow_recalc_fallback: bool = False):
        self.server_path = server_path
        self.storage_path = storage_path
        self.api_key = api_key
        self.model = model

        # Initialize components
        server_args = ["--allow-recalc-fallback"] if allow_recalc_fallback else []
        self.excel_client = ExcelMCPClient(server_path, storage_path, server_args=server_args)
        self.task_executor = ExcelTaskExecutor(self.excel_client, api_key, model, fresh_context_mode=fresh_context_mode, enhanced_excel_context=enhanced_excel_context, base_url=base_url)

        # CLI state
        self.running = True
        self.current_task = None

    def initialize(self):
        """Initialize the Excel agent"""
        print("🚀 Initializing Excel CLI Agent...")
        try:
            self.excel_client.connect()
            print("✅ Excel MCP Server connected")
            print("📁 Storage directory:", self.storage_path)
            print("🧠 AI Model:", self.model)
            print("\nType 'help' for available commands or 'quit' to exit.")
            return True
        except Exception as e:
            print(f"❌ Failed to initialize: {str(e)}")
            return False

    def cleanup(self):
        """Clean up resources"""
        self.excel_client.disconnect()

    def print_help(self):
        """Print help information"""
        help_text = """
        === EXCEL CLI AGENT COMMANDS ===

        TASK EXECUTION:
        task <description>           - Execute Excel automation task with AI

        CONTEXT:
        addcontext <pdf1> <pdf2>...  - Add PDF files to task context
        addexcel <file1> <file2>...  - Add Excel files to task context
        clearcontext                 - Clear all context (PDFs and Excel files)
        showcontext                  - Show current context files

        CONFIGURATION:
        config iterations <n>        - Set max AI iterations per task (default 30)
        config verbose <true|false>  - Enable/disable full AI response printing
        config                       - Show current config settings

        GENERAL:
        help                         - Show this help
        clear                        - Clear screen
        quit, exit                   - Exit the agent

        EXAMPLES:
        task "Create a budget spreadsheet with income and expense tracking"
        addcontext requirements.pdf data.pdf
        addexcel source_data.xlsx reference.xlsx
        task "Create analysis based on the context documents"
        config iterations 50
        config verbose true
        """
        print(help_text)

    def print_welcome(self):
        """Print welcome message"""
        welcome = """
        ╔══════════════════════════════════════════════════════════════╗
        ║                    EXCEL CLI AGENT                           ║
        ║                                                              ║
        ║  Intelligent Excel automation with iterative task execution  ║
        ║  Type 'help' for commands or describe your task naturally    ║
        ╚══════════════════════════════════════════════════════════════╝
        """
        print(welcome)

    def handle_task_command(self, args: List[str]):
        """Handle task execution command"""
        if not args:
            print("❌ Please provide a task description")
            print("Example: task \"Create a sales report with charts\"")
            return

        task_description = " ".join(args)
        print(f"🎯 Executing task: {task_description}")

        try:
            self.current_task = self.task_executor.execute_task(task_description)

            # Show summary
            summary = self.task_executor.get_execution_summary(self.current_task, show_steps=False)
            print("\n" + summary)
            if self.current_task.status == TaskStatus.COMPLETED:
                final_text = self.current_task.final_result or "Task completed"
                print(f"🧾 JOB SUMMARY: {final_text}")

            # Show next steps if task needs clarification
            if self.current_task.status == TaskStatus.NEEDS_CLARIFICATION:
                print("💭 Task paused - please provide clarification or use 'continue' to proceed")

        except Exception as e:
            print(f"💥 Task execution failed: {str(e)}")

    def handle_config_command(self, args: List[str]):
        """Handle configuration changes (e.g., iterations, verbose)"""
        if not args:
            print(f"⚙️  iterations = {self.task_executor.default_max_iterations}")
            print(f"⚙️  verbose = {self.task_executor.verbose}")
            return
        key = args[0].lower()
        if key in ["iterations", "iter", "max-iter", "max"]:
            n = None
            for tok in args[1:]:
                try:
                    n = int(tok)
                    break
                except ValueError:
                    continue
            if n is None:
                print("❌ Please specify an integer, e.g., config iterations 30")
                return
            self.task_executor.set_max_iterations(n)
            print(f"✅ Set default max iterations to {self.task_executor.default_max_iterations}")
        elif key in ["verbose", "debug", "v"]:
            if len(args) < 2:
                print("❌ Please specify true/false, e.g., config verbose true")
                return
            value = args[1].lower()
            enabled = value in ["true", "1", "yes", "on", "enable", "enabled"]
            self.task_executor.set_verbose(enabled)
            print(f"✅ Set verbose mode to {enabled}")
        else:
            print("❌ Unknown config key. Supported: iterations <n>, verbose <true|false>")

    def handle_add_context_command(self, args: List[str]):
        """Handle adding PDF context"""
        if not args:
            print("❌ Please specify PDF file paths")
            return

        result = self.task_executor.add_context_pdfs(args)

        if result['added']:
            print(f"✅ Added {len(result['added'])} PDF(s) to context:")
            for pdf in result['added']:
                print(f"   📄 {pdf}")

        if result.get('not_found'):
            print(f"❌ Could not find {len(result['not_found'])} file(s):")
            for pdf in result['not_found']:
                print(f"   ❓ {pdf}")
            print(f"💡 Files must be in storage directory: {self.task_executor.excel_client.storage_path}")

        print(f"📊 Total context PDFs: {result['total_context_pdfs']}")

    def handle_add_excel_context_command(self, args: List[str]):
        """Handle adding Excel context"""
        if not args:
            print("❌ Please specify Excel file paths")
            return

        result = self.task_executor.add_context_excels(args)

        if result['added']:
            print(f"✅ Added {len(result['added'])} Excel file(s) to context:")
            for excel in result['added']:
                print(f"   📊 {excel}")

        if result.get('not_found'):
            print(f"❌ Could not find {len(result['not_found'])} file(s):")
            for excel in result['not_found']:
                print(f"   ❓ {excel}")
            print(f"💡 Files must be in storage directory: {self.task_executor.excel_client.storage_path}")

        print(f"📊 Total context Excel files: {result['total_context_excels']}")

    def handle_clear_context_command(self):
        """Handle clearing PDF and Excel context"""
        pdf_result = self.task_executor.clear_context_pdfs()
        excel_result = self.task_executor.clear_context_excels()
        total_cleared = pdf_result['cleared'] + excel_result['cleared']
        print(f"🗑️ Cleared {pdf_result['cleared']} PDF(s) and {excel_result['cleared']} Excel file(s) from context ({total_cleared} total)")

    def handle_show_context_command(self):
        """Handle showing current PDF and Excel context"""
        has_context = False

        if self.task_executor.context_pdfs:
            has_context = True
            print(f"📋 Current context PDFs ({len(self.task_executor.context_pdfs)}):")
            for pdf in self.task_executor.context_pdfs:
                print(f"   📄 {pdf}")

        if self.task_executor.context_excels:
            has_context = True
            if self.task_executor.context_pdfs:
                print()
            print(f"📋 Current context Excel files ({len(self.task_executor.context_excels)}):")
            for excel in self.task_executor.context_excels:
                print(f"   📊 {excel}")

        if not has_context:
            print("📋 No context files loaded")

    def process_command(self, user_input: str):
        """Process a user command"""
        parts = user_input.strip().split()
        if not parts:
            return

        command = parts[0].lower()
        args = parts[1:]

        # Task execution
        if command == "task":
            self.handle_task_command(args)

        # Context commands
        elif command in ["addcontext", "add_context"]:
            self.handle_add_context_command(args)
        elif command in ["addexcel", "add_excel"]:
            self.handle_add_excel_context_command(args)
        elif command in ["clearcontext", "clear_context"]:
            self.handle_clear_context_command()
        elif command in ["showcontext", "show_context"]:
            self.handle_show_context_command()

        # Configuration
        elif command in ["config", "set"]:
            self.handle_config_command(args)

        # General commands
        elif command == "help":
            self.print_help()
        elif command == "clear":
            os.system('clear' if os.name == 'posix' else 'cls')
        elif command in ["quit", "exit"]:
            self.running = False
        else:
            print(f"❓ Unknown command: {command}")
            print("💡 Type 'help' for available commands")

    def run_interactive(self):
        """Run interactive CLI mode"""
        if not self.initialize():
            return 1

        self.print_welcome()

        try:
            while self.running:
                try:
                    user_input = input("\n🤖 excel> ").strip()
                    if not user_input:
                        continue

                    # If input doesn't start with a command, treat as task
                    parts = user_input.split()
                    if parts and parts[0].lower() not in [
                        'task', 'addcontext', 'add_context', 'addexcel', 'add_excel',
                        'clearcontext', 'clear_context', 'showcontext', 'show_context',
                        'config', 'set', 'help', 'clear', 'quit', 'exit'
                    ]:
                        # Treat as natural language task
                        self.handle_task_command(parts)
                    else:
                        self.process_command(user_input)

                except KeyboardInterrupt:
                    print("\n👋 Use 'quit' to exit")
                except EOFError:
                    print("\n👋 Goodbye!")
                    self.running = False

        finally:
            self.cleanup()

        return 0


def main():
    """Main entry point (synchronous)"""
    parser = argparse.ArgumentParser(description="Excel CLI Agent - Intelligent Excel Automation")
    import excel_mcp_server
    default_server_path = str(Path(excel_mcp_server.__file__).parent / "server.py")
    parser.add_argument("--server-path", default=default_server_path,
                       help="Path to Excel MCP server script")
    parser.add_argument("--storage-path", default=".",
                       help="Excel files storage directory")
    parser.add_argument("--api-key", help="OpenAI API key (or set OPENAI_API_KEY env var)")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="AI model to use")
    parser.add_argument("--max-iterations", type=int, default=DEFAULT_MAX_ITERATIONS, help="Maximum iterations per task")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose mode (print full AI responses)")
    parser.add_argument("--batch", nargs="+", help="Run commands in batch mode")
    parser.add_argument("--batch-config", help="Path to YAML batch configuration file: run every selected "
                                               "benchmark task from <data_root>/tasks/ and record attempts "
                                               "under <output_root>/attempts/<agent_model_name>/")
    parser.add_argument("--dry-run", action="store_true",
                        help="With --batch-config: resolve the config, the tasks and the output locations, "
                             "print the plan and exit. No MCP server, no LibreOffice, no model call, no API key needed.")
    parser.add_argument("--custom_reasoning", action="store_true",
                       help="Include reasoning field in JSON responses (uses more tokens, may hit context limit)")
    parser.add_argument("--fresh-context", action="store_true",
                       help="Enable fresh context mode: load solution.xlsx each iteration (no execution history)")
    parser.add_argument("--enhanced-excel", action="store_true", default=True,
                       help="Use enhanced grid format for Excel context files (default: True)")
    parser.add_argument("--enable-langfuse", action="store_true", default=False,
                       help="Enable Langfuse observability (disabled by default to prevent hangs)")
    parser.add_argument("--snapshot-iterations", action="store_true", default=False,
                       help="Save solution.xlsx copy + AI context text after each iteration")
    parser.add_argument("--base-url", default=None,
                       help="API base URL (auto-detects OpenAI vs Anthropic). Works with vLLM, SGLang, OpenRouter, etc.")
    parser.add_argument("--allow-recalc-fallback", action="store_true", default=False,
                       help="Interactive/--batch mode only: if LibreOffice is unavailable, run with "
                            "the limited _eval_formula fallback instead of failing at startup. "
                            "Batch-config runs use the allow_recalc_fallback config key instead.")

    args = parser.parse_args()

    # Load environment variables from .env file in current working directory
    load_dotenv(Path.cwd() / ".env")

    if args.dry_run:
        if not args.batch_config:
            parser.error("--dry-run needs --batch-config")
        return dry_run_from_config(args.batch_config, args.server_path)

    # Get API key — env vars first (flag > .env/shell), then the repo
    # config (keys.* in <MBABench>/config/config.yaml). Same OpenAI ->
    # OpenRouter -> Anthropic order in both tiers; TaskExecutor re-resolves
    # per provider from base_url, so this only needs to find *a* key.
    base_url = getattr(args, 'base_url', None) or os.getenv("BASE_URL")
    api_key = (args.api_key or os.getenv("OPENAI_API_KEY") or os.getenv("OPENROUTER_API_KEY")
               or os.getenv("ANTHROPIC_API_KEY") or os.getenv("FORGE_API_KEY")
               or repo_value("keys", "openai_api_key")
               or repo_value("keys", "openrouter_api_key") or repo_value("keys", "anthropic_api_key")
               or repo_value("keys", "forge_api_key"))
    if not api_key and not (base_url and "localhost" in base_url):
        print("❌ API key required. Set OPENAI_API_KEY, OPENROUTER_API_KEY, "
              "ANTHROPIC_API_KEY or FORGE_API_KEY in .env, or keys.* in <MBABench>/config/config.yaml")
        return 1
    if not api_key:
        api_key = "no-key"  # Local servers (vLLM/SGLang) don't need a real key

    # Batch mode: every selected benchmark task, recorded to <output_root>.
    if args.batch_config:
        print(f"🚀 Starting batch execution from config: {args.batch_config}")
        custom_reasoning = getattr(args, 'custom_reasoning', False)
        batch_result = run_batch_from_config(args.batch_config, args.server_path, api_key,
                                             custom_reasoning, args.enable_langfuse)
        return 0 if batch_result.successful == batch_result.total_workspaces else 1

    # Create storage directory
    Path(args.storage_path).mkdir(parents=True, exist_ok=True)

    # Initialize agent with new context flags
    agent = ExcelCLIAgent(
        args.server_path,
        args.storage_path,
        api_key,
        args.model,
        fresh_context_mode=args.fresh_context,
        enhanced_excel_context=args.enhanced_excel,
        base_url=args.base_url,
        allow_recalc_fallback=args.allow_recalc_fallback,
    )
    agent.task_executor.set_max_iterations(args.max_iterations)
    agent.task_executor.set_verbose(args.verbose)
    agent.task_executor.snapshot_iterations = args.snapshot_iterations

    # Run in batch or interactive mode
    if args.batch:
        if not agent.initialize():
            return 1

        for command in args.batch:
            print(f"\n🤖 Executing: {command}")
            agent.process_command(command)

        agent.cleanup()
        return 0
    else:
        return agent.run_interactive()


if __name__ == "__main__":
    exit(main())
