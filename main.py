import os
from pathlib import Path
import logging
from termcolor import colored
from SpreadsheetLLMWrapper import SpreadsheetLLMWrapper


# Custom colored formatter for stdout
class ColoredFormatter(logging.Formatter):
    """Custom formatter that adds colors to log levels"""

    LEVEL_COLORS = {
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'red',
    }

    def format(self, record):
        # Format the message with the parent formatter
        log_message = super().format(record)

        # Add color to the level name in the output
        level_name = record.levelname
        if level_name in self.LEVEL_COLORS:
            # Color the entire log line based on level
            log_message = colored(log_message, self.LEVEL_COLORS[level_name])

        return log_message


# Configure logging with colored output for stdout
# Only configure root logger to avoid duplicate logs
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
colored_formatter = ColoredFormatter(
    fmt='[%(asctime)s][%(levelname)s] %(name)s: %(message)s',
    datefmt='%H:%M:%S'
)
console_handler.setFormatter(colored_formatter)

# Configure root logger
logging.root.setLevel(logging.DEBUG)
logging.root.handlers = []
logging.root.addHandler(console_handler)

# Get logger for this module
logger = logging.getLogger(__name__)

original_size = 0
new_size = 0


if __name__ == "__main__":
    import argparse

    # Create argument parser
    parser = argparse.ArgumentParser(
        description='SpreadsheetLLM: Compress spreadsheet files using LLM-friendly format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Compress a single file in simple mode
  python main.py input.xlsx

  # Compress with format-aware mode
  python main.py input.xlsx --format-aware

  # Specify custom output directory
  python main.py input.xlsx -f -o results/
        """
    )

    # Required arguments
    parser.add_argument(
        'input_file',
        type=str,
        help='Path to the input Excel file (.xlsx)'
    )

    # Optional arguments
    parser.add_argument(
        '-f', '--format-aware',
        action='store_true',
        help='Enable format-aware aggregation (groups by value AND category)'
    )

    parser.add_argument(
        '-o', '--output-dir',
        type=str,
        default='output',
        help='Output directory for compressed files (default: output/)'
    )

    # Parse arguments
    args = parser.parse_args()

    # Log mode
    if args.format_aware:
        logger.info("Running in FORMAT-AWARE mode (dict groups by value AND category)")
    else:
        logger.info("Running in SIMPLE mode (dict groups by value only)")
        logger.info("Use --format-aware or -f flag to enable format-aware aggregation")

    # Validate input file
    file = Path(args.input_file)
    if not file.exists():
        logger.error(f"Input file not found: {file}")
        exit(1)

    if file.suffix.lower() not in ['.xlsx', '.xls']:
        logger.error(f"Input file must be an Excel file (.xlsx or .xls): {file}")
        exit(1)

    logger.info(f"Input file: {file}")

    # Create output directory if it doesn't exist
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"Output directory: {output_dir}")

    # Process spreadsheet
    wrapper = SpreadsheetLLMWrapper(format_aware=args.format_aware)

    if wb := wrapper.read_spreadsheet(file):
        if compress_result := wrapper.compress_spreadsheet(wb):
            areas, compress_dict, sheet_compressor = compress_result

            # Generate output file names
            base_name = output_dir / file.stem  # Use stem to get filename without extension
            suffix = "_format_aware" if args.format_aware else ""

            areas_file = str(base_name) + suffix + "_areas.txt"
            dict_file = str(base_name) + suffix + "_dict.txt"
            mapping_file = str(base_name) + suffix + "_mapping.json"

            # Write output files
            wrapper.write_areas(areas_file, areas, sheet_compressor)
            wrapper.write_dict(dict_file, compress_dict)
            wrapper.write_mapping(mapping_file, sheet_compressor)

            logger.info("Output files:")
            logger.info(f"  - {areas_file}")
            logger.info(f"  - {dict_file}")
            logger.info(f"  - {mapping_file}")

            # Calculate compression ratio
            original_size += os.path.getsize(file)
            new_size += os.path.getsize(areas_file)
            new_size += os.path.getsize(dict_file)
            logger.info("Compression Ratio: {:.2f}".format(original_size / new_size))
        else:
            logger.error("Compression failed")
            exit(1)
    else:
        logger.error("Failed to read spreadsheet")
        exit(1)
