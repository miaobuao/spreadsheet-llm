import asyncio
import json
import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

from benchmark.agent import run_claude_agent
from claude_agent_sdk import TextBlock
from rich.logging import RichHandler

# Create logs directory if it doesn't exist
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)

# Create a log file with timestamp
log_file = log_dir / f"dsbench_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

# Configure logging with both console (Rich) and file handlers
file_handler = logging.FileHandler(log_file, encoding="utf-8")
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(
    logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
)

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    datefmt="[%X]",
    handlers=[RichHandler(rich_tracebacks=True), file_handler],
)

logger = logging.getLogger(__name__)
logger.info(f"Logging to file: {log_file}")


def setup_workspace(source_folder: Path, workspace_root: Path, task_id: str) -> Path:
    """
    Create a workspace for a task and copy files (excluding PDFs) from source folder.

    Args:
        source_folder: The source folder containing data files
        workspace_root: Root directory for all workspaces
        task_id: Unique identifier for this task

    Returns:
        Path to the created workspace
    """
    # Create workspace directory
    workspace_path = workspace_root / task_id
    workspace_path.mkdir(parents=True, exist_ok=True)

    # Create a subfolder with the source folder name
    data_folder = workspace_path / source_folder.name
    data_folder.mkdir(exist_ok=True)

    # Copy all files except PDFs
    for file_path in source_folder.iterdir():
        if file_path.is_file() and file_path.suffix.lower() != ".pdf":
            shutil.copy2(file_path, data_folder / file_path.name)
            logger.debug(f"Copied {file_path.name} to workspace")

    return workspace_path


def parse_task_selector(selector: str) -> tuple[str, int]:
    """
    Parse a task selector string like "1-14" into (folder, question_number).

    Args:
        selector: String in format "folder_num-question_num" (e.g., "1-14")

    Returns:
        Tuple of (folder_name, question_number) (e.g., ("00000001", 14))
    """
    match = re.match(r"(\d+)-(\d+)", selector)
    if not match:
        raise ValueError(
            f"Invalid selector format: {selector}. Expected format: 'folder-question' (e.g., '1-14')"
        )

    folder_num = int(match.group(1))
    question_num = int(match.group(2))

    # Convert folder number to folder name (e.g., 1 -> "00000001")
    folder_name = f"{folder_num:08d}"

    return folder_name, question_num


def extract_question_number(question: str) -> Optional[int]:
    """
    Extract question number from the question text.

    Args:
        question: Question text (e.g., "Question 14\nHow many tonnes...")

    Returns:
        Question number if found, None otherwise
    """
    match = re.match(r"Question (\d+)", question)
    if match:
        return int(match.group(1))
    return None


def filter_dataset_by_selectors(
    dataset: list, selectors: Optional[list[str]] = None
) -> list[tuple[int, dict]]:
    """
    Filter dataset by task selectors.

    Args:
        dataset: Full dataset loaded from JSON
        selectors: List of task selectors (e.g., ["1-14", "1-8", "4-41"])
                  If None, returns all items with their indices

    Returns:
        List of (original_index, test_case) tuples
    """
    if selectors is None:
        # Return all items with their indices
        return list(enumerate(dataset))

    # Parse all selectors
    selected_tasks = set()
    for selector in selectors:
        folder_name, question_num = parse_task_selector(selector)
        selected_tasks.add((folder_name, question_num))

    logger.info(f"Filtering dataset for {len(selected_tasks)} selected tasks")

    # Filter dataset
    filtered = []
    for idx, test_case in enumerate(dataset):
        folder = test_case["folder"]
        question_num = extract_question_number(test_case["Question"])

        if question_num and (folder, question_num) in selected_tasks:
            filtered.append((idx, test_case))

    logger.info(f"Found {len(filtered)} matching test cases in dataset")

    return filtered


async def run_dsbench_tests(
    dataset_path: Path,
    base_folder: Path,
    workspace_root: Path = Path("workspaces"),
    task_selectors: Optional[list[str]] = None,
):
    """
    Run tests on the dsbench dataset.

    Args:
        dataset_path: Path to the dataset JSON file
        base_folder: Base folder containing the data folders
        workspace_root: Root directory for workspaces
        task_selectors: Optional list of task selectors (e.g., ["1-14", "4-41"])
                       Format: "folder_num-question_num" where folder_num is the folder
                       number (e.g., 1 for "00000001") and question_num is the question
                       number from the question text (e.g., 14 for "Question 14").
                       If None, runs all tests.
    """

    # Load the dataset
    with open(dataset_path, "r") as f:
        dataset = json.load(f)

    logger.info(f"Loaded {len(dataset)} test cases from {dataset_path}")

    # Filter dataset by selectors
    filtered_dataset = filter_dataset_by_selectors(dataset, task_selectors)

    if task_selectors:
        logger.info(f"Running {len(filtered_dataset)} selected test cases")
    else:
        logger.info(f"Running all {len(filtered_dataset)} test cases")

    # Create workspace root directory
    workspace_root.mkdir(exist_ok=True)

    # Save project root directory to pass to agents
    project_root = Path.cwd()

    results = []
    total_tests = len(filtered_dataset)

    for test_num, (original_idx, test_case) in enumerate(filtered_dataset, 1):
        question = test_case["Question"]
        folder = test_case["folder"]
        expected_answer = test_case["Final answer"]
        question_num = extract_question_number(question)

        source_folder = base_folder / folder
        task_id = f"task_{folder}_q{question_num}" if question_num else f"task_{folder}"

        logger.info(f"\n{'='*80}")
        logger.info(
            f"Running test {test_num}/{total_tests} (dataset index: {original_idx})"
        )
        logger.info(f"Task ID: {folder}-{question_num}")
        logger.info(f"Folder: {folder}")
        logger.info(f"Question: {question}")
        logger.info(f"Expected answer: {expected_answer}")
        logger.info(f"{'='*80}\n")

        try:
            # Setup workspace for this task
            workspace_path = setup_workspace(source_folder, workspace_root, task_id)
            logger.info(f"Created workspace at: {workspace_path}")

            # Create prompt with workspace path and data folder
            data_folder = workspace_path / source_folder.name
            prompt = f"""You need to understand the spreadsheet structure and extract data from files.

Question: {question}

**YOUR TASK:**
1. **Understand and document the spreadsheet structure in README.md**
2. **Create main.py to extract data, then run it**

Instructions:
1. The data files are located in: {data_folder}
2. First, analyze the spreadsheet files to understand their structure
3. Create a file named `README.md` in the workspace root that documents:
   - File structure and format (Excel, CSV, etc.)
   - Sheet names and their purposes
   - Column headers and their meanings
   - Data types and value ranges
   - Key observations about the data layout
   - Any relationships between different sheets/files
4. Then create `main.py` that extracts relevant raw data from the files
5. Execute `main.py` using 'uv run python main.py'
6. The script must output extracted data as JSON to stdout
7. **DO NOT analyze, calculate, or answer the question**
8. **DO NOT hardcode answers or computed values in the script**
9. Only extract raw data that directly exists in the files

Requirements for README.md:
- Clear description of each file's purpose
- Documentation of all sheets and their structure
- Explanation of column meanings
- Notes on data quality or special patterns

Requirements for main.py:
- Read the data files in {data_folder}
- Extract only raw data needed to answer the question (e.g., base values, rates, dates)
- Print a JSON object to stdout:
```json
{{
  "extracted_data": {{
    // Raw data extracted from files
  }},
  "data_sources": [
    // Files and locations where data was found
  ]
}}
```

After running the script, copy the JSON output in your final response.

Example of what to extract (NOT calculate):
- For price questions: extract base_price and inflation_rate (don't calculate final price)
- For date questions: extract date values (don't compute date differences)

Please create README.md first, then create and run main.py, and show me the JSON output."""

            # Run the agent to create and execute main.py
            result = await run_claude_agent(
                prompt=prompt,
                workspace_path=workspace_path,
                project_root=project_root,
            )

            # Extract the response text (should contain the JSON output)
            extracted_data_text = ""
            if result.response:
                for block in result.response.content:
                    if isinstance(block, TextBlock):
                        extracted_data_text += block.text

            logger.info("\n📊 Data extraction completed")
            logger.info(f"Agent response:\n{extracted_data_text}")

            # Now use AI to analyze the extracted data and answer the question
            analysis_prompt = f"""Based on the spreadsheet structure documentation and extracted data below, please answer the following question.

Question: {question}

**Context:**
1. A README.md file in the workspace documents the spreadsheet structure
2. You can read this file to understand the data layout and column meanings
3. The extracted data is provided below

README.md location: {workspace_path}/README.md

Extracted Data:
{extracted_data_text}

Instructions:
1. First, read the README.md file to understand the spreadsheet structure
2. Parse and analyze the extracted data carefully using the context from README.md
3. Perform any necessary calculations or reasoning
4. Explain your reasoning process
5. **IMPORTANT**: Output your final answer on the LAST LINE of your response
6. For multiple choice questions, the last line should ONLY contain the letter (e.g., "A", "B", "C")
7. For numeric answers, the last line should ONLY contain the exact number

Format your response like this:
[Read README.md and understand the structure...]
[Your analysis and reasoning here...]

Final Answer: [Your answer here - just the letter or number]

Please provide your answer now."""

            logger.info(f"\n🤔 Analyzing data to answer question...")

            # Run AI analysis (second agent call)
            analysis_result = await run_claude_agent(
                prompt=analysis_prompt,
                workspace_path=workspace_path,
                project_root=project_root,
            )

            # Extract the final answer from analysis (last line only)
            full_response = ""
            if analysis_result.response:
                for block in analysis_result.response.content:
                    if isinstance(block, TextBlock):
                        full_response += block.text

            # Extract answer from the last line
            lines = full_response.strip().split("\n")
            last_line = lines[-1].strip() if lines else ""

            # Try to extract answer after "Final Answer:" prefix
            if "final answer:" in last_line.lower():
                final_answer = last_line.split(":", 1)[-1].strip()
            else:
                final_answer = last_line

            logger.info(f"\n💡 Full response: {full_response}")
            logger.info(f"💡 Extracted final answer from last line: {final_answer}")

            # Verify answer - check if expected answer matches the final answer line
            is_correct = expected_answer.strip() in final_answer

            # Store result
            test_result = {
                "index": original_idx,
                "folder": folder,
                "question_number": question_num,
                "task_id": task_id,
                "workspace_path": str(workspace_path),
                "question": question,
                "expected_answer": expected_answer,
                "extracted_data": extracted_data_text,
                "full_response": full_response,
                "final_answer": final_answer,
                "is_correct": is_correct,
                "extraction_time": result.time,
                "analysis_time": analysis_result.time,
                "total_time": result.time + analysis_result.time,
                "extraction_usages": result.usages,
                "analysis_usages": analysis_result.usages,
            }

            results.append(test_result)

            logger.info(f"\n Test {test_num} completed in {result.time:.2f}s")
            status = "✓" if is_correct else "✗"
            logger.info(f"Result: {status} | Answer: {expected_answer}")

        except Exception as e:
            logger.error(f" Test {test_num} failed with error: {e}")
            results.append(
                {
                    "index": original_idx,
                    "folder": folder,
                    "question_number": question_num,
                    "task_id": task_id,
                    "question": question,
                    "expected_answer": expected_answer,
                    "error": str(e),
                }
            )

    # Calculate statistics
    total = len(results)
    correct = sum(1 for r in results if r.get("is_correct", False))

    # Save results with timestamp and task info
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if task_selectors:
        task_count = len(task_selectors)
        output_filename = f"dsbench_results_{timestamp}_selected_{task_count}.json"
    else:
        output_filename = f"dsbench_results_{timestamp}_all.json"

    output_path = Path(output_filename)
    with open(output_path, "w") as f:
        json.dump(
            {
                "task_selectors": task_selectors,
                "total_tests": total,
                "correct": correct,
                "accuracy": correct / total if total > 0 else 0,
                "results": results,
            },
            f,
            indent=2,
            ensure_ascii=False,
        )

    logger.info("\n" + "=" * 80)
    logger.info(f"Completed: {correct}/{total} correct ({correct/total*100:.1f}%)")
    logger.info(f"Results saved to {output_path}")
    logger.info("=" * 80)

    return results


if __name__ == "__main__":
    # Example usage:
    # Run all tests:
    #   python -m benchmark.run_dsbench
    #
    # Run specific tests:
    #   Modify task_selectors list below with desired tasks
    #   Format: "folder_num-question_num" (e.g., "1-14" for folder 00000001, Question 14)

    # Set to None to run all tests, or provide a list of task selectors
    task_selectors = [
        # "1-14", # ✅
        # "4-42", # ✅
        # "4-47", # ❌
        # "4-50",  # ❌
        # "9-13",  # ❌
        # "10-14",  # ❌
        # "10-16", # ❌
        # "13-24",  # ❌
        # "16-2",  # ❌
        # "20-33",  # ✅
        # "25-38",  # ✅
    ]

    asyncio.run(
        run_dsbench_tests(
            Path(
                "/Volumes/Yang/dev/github/spreadsheet-agent/dataset/dsbench/sample_60_cleaned.json"
            ),
            Path("/Volumes/Yang/dev/github/spreadsheet-agent/dataset/dsbench"),
            task_selectors=task_selectors,
        )
    )
