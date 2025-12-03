"""
List all sheet names in an Excel workbook.

Usage:
    uv run tools/list-sheets.py <excel_file>
"""

import argparse
import sys
from pathlib import Path

from spreadsheet_llm import SpreadsheetLLMWrapper


def main():
    parser = argparse.ArgumentParser(
        description="List all sheet names in an Excel workbook",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # List all sheets in a workbook
  uv run tools/list-sheets.py data.xlsx

  # With absolute path
  uv run tools/list-sheets.py /path/to/workbook.xlsx
        """,
    )

    parser.add_argument(
        "input_file",
        type=str,
        help="Path to the input Excel file (.xlsx, .xlsb, .xls)",
    )

    args = parser.parse_args()

    # Validate input file
    file_path = Path(args.input_file)
    if not file_path.exists():
        print(f"Error: File not found: {file_path}", file=sys.stderr)
        sys.exit(1)

    if file_path.suffix.lower() not in [".xlsx", ".xls", ".xlsb"]:
        print(
            f"Error: Input file must be an Excel file (.xlsx, .xlsb, or .xls): {file_path}",
            file=sys.stderr,
        )
        sys.exit(1)

    # Read workbook and list sheets
    wrapper = SpreadsheetLLMWrapper()
    wb = wrapper.read_spreadsheet(file_path)

    if wb is None:
        print(f"Error: Failed to read workbook: {file_path}", file=sys.stderr)
        sys.exit(1)

    # Print sheet names
    print(f"Workbook: {file_path.name}")
    print(f"Total sheets: {len(wb.sheetnames)}")
    print()
    print("Sheet names:")
    for i, sheet_name in enumerate(wb.sheetnames):
        print(f"  [{i}] {sheet_name}")


if __name__ == "__main__":
    main()
