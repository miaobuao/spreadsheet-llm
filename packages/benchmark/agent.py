import json
import logging
import platform
import time
from datetime import datetime
from pathlib import Path
from typing import NamedTuple, TypedDict

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

logger = logging.getLogger(__name__)


class ServerToolUse(TypedDict):
    web_search_requests: int
    web_fetch_requests: int


class CacheCreation(TypedDict):
    ephemeral_1h_input_tokens: int
    ephemeral_5m_input_tokens: int


class Usage(TypedDict):
    input_tokens: int
    cache_creation_input_tokens: int
    cache_read_input_tokens: int
    output_tokens: int
    server_tool_use: ServerToolUse
    service_tier: str
    cache_creation: CacheCreation


def get_system_prompt(cwd: Path | str, workspace_path: Path | str, today: str) -> str:
    return f"""You are an AI coding assistant.

<environment>
- Current Working Directory: {cwd}
- Task Workspace Path: {workspace_path}
- Platform: {platform.system()}
- Python Version: {platform.python_version()}
- Package Manager: uv
- Today's Date: {today}
</environment>

Important:
- Always use 'uv run python' to execute Python scripts
- All file operations (Read, Write, Edit) must be performed within the task workspace: {workspace_path}
- When using 'cd' in Bash commands, use subshell or always cd back to maintain working directory:
  - Use: (cd {workspace_path} && command)  # subshell, doesn't change parent cwd
  - Or use: cd {workspace_path} && command && cd {cwd}  # explicit cd back
- Skills and CLI tools always execute from {cwd}, so they don't need cd
- Example: Write main.py in workspace, then run: (cd {workspace_path} && uv run python main.py)"""


class RunClaudeAgentResult(NamedTuple):
    response: AssistantMessage | None
    usages: list[Usage]
    time: float


async def run_claude_agent(
    client: ClaudeSDKClient | None = None,
    prompt: str = "",
    workspace_path: Path | str | None = None,
    project_root: Path | str | None = None,
):
    start_time = time.time()

    # Get project root (passed in or use current directory)
    if project_root is None:
        project_root = Path.cwd()

    # Save original working directory to restore later
    original_cwd = Path.cwd()

    # Create client if not provided
    should_disconnect = False
    if client is None:
        if workspace_path is None:
            raise ValueError("workspace_path is required when client is not provided")

        client = ClaudeSDKClient(
            ClaudeAgentOptions(
                setting_sources=["project"],
                allowed_tools=["Read", "Write", "Edit", "Bash", "Skill"],
                permission_mode="acceptEdits",
                system_prompt=get_system_prompt(
                    cwd=project_root,
                    workspace_path=workspace_path,
                    today=datetime.now().strftime("%Y-%m-%d"),
                ),
            )
        )
        await client.connect()
        should_disconnect = True

    await client.query(prompt=prompt)

    last_response = None
    usages = []
    async for msg in client.receive_response():
        if hasattr(msg, "usage"):
            usages.append(msg.usage)  # type: ignore
        if isinstance(msg, SystemMessage):
            logger.debug(f"System message: {msg}")
            continue
        elif isinstance(msg, AssistantMessage):
            last_response = msg
            for block in msg.content:
                if isinstance(block, TextBlock):
                    logger.info(f"Text response: {block.text}")
                elif isinstance(block, ThinkingBlock):
                    logger.debug(f"Thinking: {block.thinking}")
                elif isinstance(block, ToolUseBlock):
                    params_str = json.dumps(block.input, indent=2, ensure_ascii=False)
                    logger.info(
                        f"Tool Use - Name: {block.name}, ID: {block.id}\nParameters:\n{params_str}"
                    )
                elif isinstance(block, ToolResultBlock):
                    status = "SUCCESS" if not block.is_error else "ERROR"
                    logger.info(
                        f"Tool Result [{status}] - Tool ID: {block.tool_use_id}"
                    )
                else:
                    logger.debug(f"Unknown block type: {block}")

    elapsed_time = time.time() - start_time
    logger.info("Claude Agent processing completed")

    # Disconnect client if we created it
    if should_disconnect:
        await client.disconnect()

    # Restore original working directory
    import os

    os.chdir(original_cwd)

    return RunClaudeAgentResult(
        response=last_response, usages=usages, time=elapsed_time
    )
