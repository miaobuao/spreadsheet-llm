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

# Recent Updates

- ✅ Support for both .xlsx and .xls file formats (using openpyxl)
- ✅ Format-aware mode: Types wrapped in `${}` to distinguish from literals
