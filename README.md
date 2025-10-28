![image](https://github.com/user-attachments/assets/72cabb2a-10b6-4c43-a265-bb92ec474d54)

# SpreadsheetLLM

My unofficial implementation of Microsoft's SpreadsheetLLM paper, found here: https://arxiv.org/pdf/2407.09025.

# Requirements

All requirements are listed in requirements.txt. I have attached two Dockerfiles, one for the command line utility and one for the chatbot.

You will also need to download the VFUSE dataset from TableSense, found here: https://figshare.com/projects/Versioned_Spreadsheet_Corpora/20116

Environment Variables: OPENAI_API_KEY for GPT 3.5/4, HUGGING_FACE_KEY for Llama-2/3, Phi-3, and Mistral

# Usage

## Command Line Interface

### Basic Usage

```bash
# Compress a single file (simple mode)
python main.py input.xlsx

# Use format-aware mode (types wrapped in ${})
python main.py input.xlsx --format-aware
# Or use short option
python main.py input.xlsx -f

# Specify custom output directory
python main.py input.xlsx -o results/

# Show help
python main.py --help
```

### Command Line Arguments

| Argument         | Short | Required | Description                                       |
| ---------------- | ----- | -------- | ------------------------------------------------- |
| `input_file`     | -     | ✅       | Path to input Excel file (.xlsx or .xls)          |
| `--format-aware` | `-f`  | ❌       | Enable format-aware mode (types wrapped in `${}`) |
| `--output-dir`   | `-o`  | ❌       | Output directory (default: output/)               |
| `--help`         | `-h`  | ❌       | Show help message                                 |

### Output Files

Three files are generated after compression:

1. **`*_dict.txt`** - Inverted index dictionary

   - Simple mode: `Eagles|B12,B39,B44`
   - Format-aware mode: `${Integer}|G14:G15,H15` (types wrapped in `${}`)

2. **`*_areas.txt`** - Data areas (contiguous cells of same type)

3. **`*_mapping.json`** - Coordinate mapping (compressed → original)

### Mode Comparison

**Simple Mode**: Groups by value only, preserves all original values

```bash
python main.py input.xlsx
```

**Format-Aware Mode**: Smart aggregation - "Other" type uses values, data types use categories

- Advantages: More compact dictionary, higher compression ratio
- Types wrapped in `${}` for easy distinction

```bash
python main.py input.xlsx -f
```

# Limitations

1. Only text was considered for the structural anchor-based extraction, formatting (border, color, etc.) was not considered
2. NFS Identification currently relies on regular expressions and may not be robust

# Development & Testing

## Running Tests

This project uses `pytest` for testing. Tests are located in the `tests/` directory.

### Prerequisites

```bash
# Install test dependencies (if not already installed)
pip install pytest pytest-cov
```

### Basic Test Commands

```bash
# Run all tests
pytest tests/ -v

# Run tests for a specific module
pytest tests/test_cell_range_utils.py -v

# Run tests with coverage report
pytest tests/ --cov=spreadsheet_llm --cov-report=html

# Show print statements and performance output
pytest tests/ -s

# Run only performance tests
pytest tests/ -k "performance" -v
```

### Test Organization

#### `tests/test_cell_range_utils.py` (50 tests)
Comprehensive tests for Excel range utilities:
- **TestColToIndex** - Column letter to index conversion (5 tests)
- **TestIndexToCol** - Index to column letter conversion (6 tests)
- **TestParseExcelRange** - Parse range strings like "A1:D10" (7 tests)
- **TestBoxToRange** - Convert box coordinates to ranges (6 tests)
- **TestCombineCells** - Cell range combination algorithm
  - Basic functionality (7 tests)
  - Edge cases (7 tests)
  - Precision requirements (4 tests)
  - Performance benchmarks (4 tests)
  - Real-world scenarios (4 tests)

#### Key Test Features
- ✅ **100% precision** - No false range merging for sparse data
- ✅ **High performance** - <1ms for 500 cells
- ✅ **Edge case coverage** - L-shapes, T-shapes, gaps, scattered cells
- ✅ **Round-trip validation** - Parse → Convert → Parse consistency

### Advanced Testing

```bash
# Filter tests by keyword
pytest tests/ -k "parse or box" -v

# Run only failed tests from last run
pytest --lf

# Run tests in parallel (requires pytest-xdist)
pip install pytest-xdist
pytest tests/ -n auto

# Generate detailed coverage report
pytest tests/ --cov=spreadsheet_llm --cov-report=term-missing

# Stop on first failure
pytest tests/ -x
```

### Test Configuration

The project uses `pytest.ini` for configuration and `tests/conftest.py` for shared fixtures and setup.

## Project Structure

```
spreadsheet-llm-unofficial/
├── src/
│   └── spreadsheet_llm/
│       ├── cell_range_utils.py    # Excel range utilities
│       ├── sheet_compressor.py    # Compression algorithms
│       └── spreadsheet_llm_wrapper.py
├── tests/
│   ├── conftest.py                    # Pytest configuration
│   ├── test_cell_range_utils.py       # 50 tests for range utils
│   ├── test_tablesense.py             # TableSense evaluation
│   └── test_dsbench_recognition.py    # DSBench batch processing
├── pytest.ini                         # Test settings
└── main.py                            # CLI entry point
```

## Batch Processing (DSBench)

### Overview

The `test_dsbench_recognition.py` script processes multiple Excel files from a directory, identifies table regions using LLM-based recognition, and outputs structured results.

### Features

- 🔄 **Batch processing** - Process entire directories of Excel files
- 📊 **Rich logging** - Beautiful console output with progress bars
- 🚀 **Smart caching** - Skip already processed files
- 📝 **JSON output** - Structured recognition results
- 🎯 **LLM-powered** - Intelligent table region identification

### Usage

```bash
# Set OpenAI API key (required)
export OPENAI_API_KEY=your-api-key-here

# Optional: Choose a different model (default: gpt-4o-mini)
export MODEL_NAME=gpt-4o

# Run the script
python tests/test_dsbench_recognition.py
```

### Configuration

Edit the script to change default paths:

```python
# Source directory containing Excel files
DSBENCH_DIR = Path("/path/to/your/excel/files")

# Output directory for results
OUTPUT_DIR = Path("output/dsbench_recognition")
```

### Output Files

The script generates:

1. **`recognition_results_<timestamp>.json`** - Detailed recognition results
   ```json
   {
     "metadata": {
       "timestamp": "2025-01-28T10:30:00",
       "total_files": 35,
       "successful": 33,
       "model": "gpt-4o-mini"
     },
     "results": [
       {
         "file_name": "example.xlsx",
         "compression_stats": { ... },
         "recognition": {
           "num_regions": 3,
           "regions": [
             {"title": "Sales Data", "range": "A1:F20"},
             {"title": "Summary", "range": "H1:J10"}
           ]
         }
       }
     ]
   }
   ```

2. **`recognition_cache.json`** - Cache to skip processed files

3. **`logs/dsbench_recognition_<timestamp>.log`** - Detailed logs

### Example Output

```
DSBench Table Recognition Pipeline
Source: /path/to/dsbench
Output: output/dsbench_recognition

✓ Processing: example.xlsx
✓ Found 3 regions in example.xlsx

Processing Complete!
  • Processed: 10 files
  • Cached: 23 files
  • Failed: 2 files
  • Total: 33 successful

┌─────────────────────┬───────────────┬─────────┬─────────┐
│ File                │ Original Size │ Anchors │ Regions │
├─────────────────────┼───────────────┼─────────┼─────────┤
│ example.xlsx        │ 100×50        │ 30×15   │ 3       │
│ data.xlsx           │ 200×80        │ 45×20   │ 5       │
└─────────────────────┴───────────────┴─────────┴─────────┘

Statistics:
  • Total regions identified: 98
  • Average per file: 3.0
```

### Features Details

- **Automatic caching** - Processed files are cached; rerun the script to process only new files
- **Progress tracking** - Real-time progress bar with file counts
- **Error handling** - Failed files are logged but don't stop processing
- **Rich console** - Color-coded output with tables and progress bars
- **Flexible logging** - Console shows INFO+, file logs DEBUG+ for troubleshooting

# Recent Updates

- ✅ Support for both .xlsx and .xls file formats (using openpyxl)
- ✅ Format-aware mode: Types wrapped in `${}` to distinguish from literals
