---
name: list-sheets
description: List all sheet names in an Excel workbook. Use this when users need to know what sheets are available in a workbook before processing.
---

# List Sheets CLI Tool

## When to Use This Skill

Use this tool when the user asks to:
- See what sheets are in an Excel workbook
- List all available worksheets in a file
- Check sheet names before processing specific sheets
- Explore the structure of a workbook
- Find the sheet index for use with other tools

## Command Syntax

```bash
uv run tools/list-sheets.py INPUT_FILE
```

## Parameters

- `INPUT_FILE` (required): Path to the Excel file (.xlsx, .xlsb, .xls)

## Output Format

The tool displays:
1. Workbook filename
2. Total number of sheets
3. List of sheet names with their indices (0-based)

Example output:
```
Workbook: sales_data.xlsx
Total sheets: 3

Sheet names:
  [0] Summary
  [1] Q1 Sales
  [2] Q2 Sales
```

## Common Usage Patterns

### 1. List Sheets in a Workbook
```bash
# Simple usage
uv run tools/list-sheets.py data.xlsx

# With absolute path
uv run tools/list-sheets.py /path/to/workbook.xlsx
```

### 2. Before Processing with spreadsheet-llm-cli
```bash
# First, list available sheets
uv run tools/list-sheets.py annual_report.xlsx

# Then process a specific sheet by index or name
uv run tools/spreadsheet-llm-cli.py annual_report.xlsx -s 1
uv run tools/spreadsheet-llm-cli.py annual_report.xlsx -s "Budget"
```

## Example Workflows

### User: "What sheets are in this Excel file?"
```bash
uv run tools/list-sheets.py financial_report.xlsx
```

### User: "Show me all the worksheets in my workbook"
```bash
uv run tools/list-sheets.py workbook.xlsx
```

### User: "I want to process the second sheet, what's it called?"
```bash
# First list sheets to see what's available
uv run tools/list-sheets.py data.xlsx

# Output shows:
#   [0] Overview
#   [1] Detailed Data
#   [2] Charts

# Then use the index or name with other tools
uv run tools/spreadsheet-llm-cli.py data.xlsx -s 1
# or
uv run tools/spreadsheet-llm-cli.py data.xlsx -s "Detailed Data"
```

## Integration with Other Tools

This tool is commonly used as a first step before:
- **spreadsheet-llm-cli**: To know which sheet to process with `-s` flag
- **Data analysis**: To understand workbook structure
- **Batch processing**: To iterate over multiple sheets

## Supported Formats

- `.xlsx` - Excel 2007+ format (most common)
- `.xlsb` - Excel binary format (faster for large files)
- `.xls` - Legacy Excel format (97-2003)

## Error Handling

The tool will display clear error messages for:
- File not found
- Invalid file format (not an Excel file)
- Corrupted workbooks that cannot be read

## Technical Notes

- Sheet indices are 0-based (first sheet is index 0)
- The tool uses the SpreadsheetLLM unified workbook interface
- Works with all Excel formats supported by the main package
- Very fast - only loads workbook metadata, not sheet contents

## Related Files

- CLI script: `tools/list-sheets.py`
- Main package: `packages/spreadsheet_llm/`
- Related tool: `tools/spreadsheet-llm-cli.py`
