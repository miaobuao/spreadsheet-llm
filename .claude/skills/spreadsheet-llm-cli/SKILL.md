---
name: spreadsheet-llm
description: Compress Excel files into LLM-friendly format and perform AI-based cell range recognition. Use this when users need to process, analyze, or extract data from spreadsheet files.
---

# SpreadsheetLLM CLI Tool

## When to Use This Skill

Use this tool when the user asks to:

- Compress or preprocess Excel/spreadsheet files for LLM analysis
- Extract tables or data ranges from spreadsheets
- Analyze spreadsheet structure and identify meaningful data regions
- Prepare large spreadsheets for token-efficient LLM processing
- Find specific data patterns in Excel files (sales tables, financial data, etc.)

## Command Syntax

```bash
uv run tools/spreadsheet-llm-cli.py [OPTIONS] INPUT_FILE
```

## Key Options

### Compression Modes

- **Simple mode** (default): Groups cells by value only
- **Format-aware mode** (`-f` or `--format-aware`): Groups by value AND formatting

### LLM Recognition

- `-r, --recognize`: Enable AI-based cell range identification (auto-detects all data regions)
- `-m, --model MODEL`: Specify model (default: `google/gemini-2.5-pro`)
  - **Recommended**: `google/gemini-2.5-pro` - Best performance and cost-effectiveness
  - Small sheets (<50 anchors): `google/gemini-2.5-pro` or `gpt-4o-mini`
  - Medium sheets (50-200): `google/gemini-2.5-pro` (recommended)
  - Large sheets (>200): `google/gemini-2.5-pro` (best value)
- `--original-coords`: **Return original spreadsheet coordinates - Recommended for agent code generation**

### Other Options

- `-o, --output-dir DIR`: Output directory (default: `output/`)
- `-s, --sheet SHEET`: Process specific sheet (by index or name)

## Common Usage Patterns

### 1. Basic Compression

```bash
# Simple compression
uv run tools/spreadsheet-llm-cli.py input.xlsx

# Format-aware compression
uv run tools/spreadsheet-llm-cli.py input.xlsx -f
```

### 2. AI-Powered Region Recognition

```bash
# Use specific model
uv run tools/spreadsheet-llm-cli.py complex.xlsx -r --original-coords -m google/gemini-2.5-pro

# Format-aware recognition for less tokens
uv run tools/spreadsheet-llm-cli.py sales.xlsx -f -r --original-coords -m google/gemini-2.5-pro
```

### 3. Process Specific Sheets

```bash
# By index (0-based)
uv run tools/spreadsheet-llm-cli.py workbook.xlsx -s 1

# By name
uv run tools/spreadsheet-llm-cli.py workbook.xlsx -s "Q4 Results"
```

## Output Files

For input file `example.xlsx` with sheet named "Sheet1", generates:

- `example_Sheet1_areas.txt`: Compressed spreadsheet representation
- `example_Sheet1_dict.txt`: Value-to-cell coordinate mappings
- `example_Sheet1_mapping.json`: Compression metadata and anchors
- `example_Sheet1_compressed.xlsx`: Compressed Excel file
- `example_Sheet1_recognition.txt`: AI recognition results (if `-r` used)

Files include sheet name in filename to distinguish different sheets. Add `_format_aware` suffix when using `-f` flag (e.g., `example_Sheet1_format_aware_areas.txt`).

## Environment Setup

### Required for Recognition

```bash
export OPENAI_API_KEY="your-api-key"
```

### Optional: Custom API Endpoint

```bash
export OPENAI_BASE_URL="http://localhost:1234/v1"  # For local LLMs
```

## Example Workflows

### User: "Compress this financial spreadsheet"

```bash
uv run tools/spreadsheet-llm-cli.py financial_report.xlsx -f -o results/
```

### User: "Find all sales tables in this Excel file"

```bash
uv run tools/spreadsheet-llm-cli.py sales_data.xlsx -r --original-coords -p "Identify all sales tables with product and revenue columns"
```

### User: "Extract data from the Budget sheet"

```bash
uv run tools/spreadsheet-llm-cli.py annual_report.xlsx -s "Budget" -r --original-coords
```

## Understanding Output

### Compression Info

The tool displays:

```
ANCHOR INFORMATION:
  Row anchors: 45 (from 500 original rows)
  Column anchors: 12 (from 50 original columns)
  Compression ratio: 9.0% rows, 24.0% columns retained
```

### Recognition Results

Shows:

- **Reasoning**: Why certain ranges were identified
- **Cell Ranges**: List of meaningful data regions with:
  - Title/description
  - Cell range coordinates (e.g., `A1:C10`)
  - Compressed encoding (for efficient LLM communication)

## Troubleshooting

### Import Error

If `spreadsheet_llm` module not found:

```bash
cd /Volumes/Yang/dev/github/spreadsheet-agent
pip install -e .
```

### Recognition Issues

- Verify `OPENAI_API_KEY` is set: `echo $OPENAI_API_KEY`
- Check model name is valid
- Ensure network connectivity to API

### Performance Tips

- Use `-m google/gemini-2.5-pro` for best performance and value (recommended)
- Use `-m gpt-4o-mini` for faster processing with lower accuracy
- Process specific sheets with `-s` for large workbooks
- Omit `-f` flag for faster compression (if formatting not important)

## Technical Notes

- Supports Excel formats: `.xlsx`, `.xlsb`, `.xls`
- Compression reduces token usage by 5-10x
- Recognition quality improves with format-aware mode (`-f`)
- Original coordinates can be preserved with `--original-coords`

## Related Files

- CLI script: `tools/spreadsheet-llm-cli.py`
- Main package: `packages/spreadsheet_llm/`
- Configuration: `pyproject.toml`
