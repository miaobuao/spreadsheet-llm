import asyncio
import json
import logging
import platform
from datetime import datetime
from pathlib import Path

from claude_agent_sdk import (
    AssistantMessage,
    ClaudeAgentOptions,
    ClaudeSDKClient,
    SystemMessage,
    TextBlock,
    ThinkingBlock,
    ToolResultBlock,
    ToolUseBlock,
)
from prompt_toolkit import PromptSession
from prompt_toolkit.formatted_text import HTML
from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel

# Initialize console and prompt session
console = Console()
session = PromptSession()

# Setup logging to file only
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file_path = log_dir / f"agent_{timestamp}.log"

logger = logging.getLogger("agent")
logger.setLevel(logging.DEBUG)

file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
file_handler.setLevel(logging.DEBUG)
file_formatter = logging.Formatter(
    "%(asctime)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
)
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

console.print(f"[dim]Logging to: {log_file_path}[/dim]")


def get_system_prompt() -> str:
    """Generate system prompt with environment information."""
    cwd = Path.cwd()
    today = datetime.now().strftime("%Y-%m-%d")

    return f"""You are an AI coding assistant that can help with various tasks.

<environment>
- Working Directory: {cwd}
- Platform: {platform.system()}
- Python Version: {platform.python_version()}
- Package Manager: uv
- Today's Date: {today}
</environment>

Important:
- Always use 'uv run' to execute Python scripts (e.g., 'uv run python script.py' or 'uv run pytest')
- Never use 'python' or 'pip' commands directly
- Always consider the current working directory and environment when working with files."""


# Initialize Claude Agent with built-in tools
client = ClaudeSDKClient(
    ClaudeAgentOptions(
        setting_sources=["project"],
        allowed_tools=["Read", "Write", "Edit", "Bash", "Skill"],
        permission_mode="acceptEdits",
        system_prompt=get_system_prompt(),
    )
)


async def main():
    console.print(
        Panel.fit(
            "[bold cyan]Claude Agent[/bold cyan]\n[dim]Type 'quit', 'exit', or 'q' to exit[/dim]",
            border_style="cyan",
        )
    )

    await client.connect()

    while True:
        console.print()
        user_input = await session.prompt_async(
            HTML("<ansigreen><b>></b></ansigreen> ")
        )

        if not user_input:
            continue

        if user_input.lower() in ["quit", "exit", "q"]:
            console.print("\n[bold yellow]👋 Goodbye![/bold yellow]")
            logger.info("=" * 80)
            logger.info("User exited the agent")
            logger.info("=" * 80)
            break

        # Log user input with separator
        logger.info("")
        logger.info("-" * 80)
        logger.info(f"USER INPUT: {user_input}")
        logger.info("-" * 80)

        await client.query(user_input)
        logger.debug("Query sent to Claude SDK")

        with console.status("[bold blue]Processing...", spinner="dots"):
            async for msg in client.receive_response():
                if isinstance(msg, SystemMessage):
                    logger.debug("System message received (skipped)")
                    continue
                elif isinstance(msg, AssistantMessage):
                    console.print()
                    logger.info(
                        f"Assistant message received with {len(msg.content)} content blocks"
                    )

                    for i, block in enumerate(msg.content, 1):
                        if isinstance(block, TextBlock):
                            console.print(
                                Panel(
                                    Markdown(block.text),
                                    border_style="blue",
                                    padding=(1, 2),
                                )
                            )
                            logger.info(f"[Block {i}] TEXT:")
                            logger.info(block.text)

                        elif isinstance(block, ThinkingBlock):
                            console.print(
                                Panel(
                                    f"[dim italic]{block.thinking}[/dim italic]",
                                    title="[dim]Thinking[/dim]",
                                    border_style="dim",
                                    padding=(0, 1),
                                )
                            )
                            logger.info(f"[Block {i}] THINKING:")
                            logger.info(block.thinking)

                        elif isinstance(block, ToolUseBlock):
                            params_str = json.dumps(
                                block.input, indent=2, ensure_ascii=False
                            )
                            console.print(
                                Panel(
                                    f"[cyan]Tool:[/cyan] {block.name}\n[dim]ID:[/dim] {block.id}\n[dim]Parameters:[/dim]\n{params_str}",
                                    title="[cyan]🔧 Tool Use[/cyan]",
                                    border_style="cyan",
                                    padding=(0, 1),
                                )
                            )
                            logger.info(f"[Block {i}] TOOL USE:")
                            logger.info(f"  Tool Name: {block.name}")
                            logger.info(f"  Tool ID: {block.id}")
                            logger.info("  Parameters:")
                            for line in params_str.split("\n"):
                                logger.info(f"    {line}")

                        elif isinstance(block, ToolResultBlock):
                            status_color = "green" if not block.is_error else "red"
                            status_icon = "✓" if not block.is_error else "✗"
                            console.print(
                                Panel(
                                    f"[{status_color}]{status_icon}[/{status_color}] [dim]Tool ID:[/dim] {block.tool_use_id}",
                                    title=f"[{status_color}]Tool Result[/{status_color}]",
                                    border_style=status_color,
                                    padding=(0, 1),
                                )
                            )
                            status_text = "SUCCESS" if not block.is_error else "ERROR"
                            logger.info(f"[Block {i}] TOOL RESULT: {status_text}")
                            logger.info(f"  Tool Use ID: {block.tool_use_id}")
                            if block.is_error:
                                logger.error(f"  Error details: {block}")

                        else:
                            console.print(block)
                            logger.warning(
                                f"[Block {i}] UNKNOWN BLOCK TYPE: {type(block)}"
                            )
                            logger.debug(f"  Block content: {block}")


if __name__ == "__main__":
    asyncio.run(main())
