import logging
import os
from pathlib import Path

from termcolor import colored

from spreadsheet_llm import SpreadsheetLLMWrapper


# Custom colored formatter for stdout
class ColoredFormatter(logging.Formatter):
    """Custom formatter that adds colors to log levels"""

    LEVEL_COLORS = {
        "DEBUG": "cyan",
        "INFO": "green",
        "WARNING": "yellow",
        "ERROR": "red",
        "CRITICAL": "red",
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
    fmt="[%(asctime)s][%(levelname)s] %(name)s: %(message)s", datefmt="%H:%M:%S"
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
        description="SpreadsheetLLM: Compress spreadsheet files using LLM-friendly format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Compress a single file in simple mode
  python main.py input.xlsx

  # Compress with format-aware mode
  python main.py input.xlsx --format-aware

  # Specify custom output directory
  python main.py input.xlsx -f -o results/

  # Enable LLM recognition with default model
  python main.py input.xlsx -r

  # Use specific model for recognition
  python main.py input.xlsx -r -m gpt-4o

  # Recognition with original coordinates
  python main.py input.xlsx -r --original-coords

  # Custom prompt for recognition
  python main.py input.xlsx -r -p "Find all tables with revenue data"

  # Use custom OpenAI-compatible API (e.g., local LLM)
  export OPENAI_BASE_URL=http://localhost:1234/v1
  export OPENAI_API_KEY=your-api-key
  python main.py input.xlsx -r

Environment Variables:
  OPENAI_API_KEY      OpenAI API key (optional if already configured)
  OPENAI_BASE_URL     Custom OpenAI-compatible API endpoint (e.g., for local models)
        """,
    )

    # Required arguments
    parser.add_argument(
        "input_file", type=str, help="Path to the input Excel file (.xlsx)"
    )

    # Optional arguments
    parser.add_argument(
        "-f",
        "--format-aware",
        action="store_true",
        help="Enable format-aware aggregation (groups by value AND category)",
    )

    parser.add_argument(
        "-o",
        "--output-dir",
        type=str,
        default="output",
        help="Output directory for compressed files (default: output/)",
    )

    parser.add_argument(
        "-r",
        "--recognize",
        action="store_true",
        help="Enable LLM-based cell range recognition (requires --model)",
    )

    parser.add_argument(
        "-m",
        "--model",
        type=str,
        default="gpt-4o-mini",
        help="LLM model to use for recognition (default: gpt-4o-mini). Examples: gpt-4o-mini, gpt-4o, gpt-3.5-turbo",
    )

    parser.add_argument(
        "-p",
        "--user-prompt",
        type=str,
        default=None,
        help="Custom prompt for LLM recognition (optional)",
    )

    parser.add_argument(
        "--original-coords",
        action="store_true",
        help="Return original spreadsheet coordinates instead of compressed coordinates (only with --recognize)",
    )

    # Parse arguments
    args = parser.parse_args()

    if args.original_coords and not args.recognize:
        logger.warning("--original-coords flag is only used with --recognize, ignoring")

    # Log mode
    if args.format_aware:
        logger.info("Running in FORMAT-AWARE mode (dict groups by value AND category)")
    else:
        logger.info("Running in SIMPLE mode (dict groups by value only)")
        logger.info("Use --format-aware or -f flag to enable format-aware aggregation")

    if args.recognize:
        logger.info(f"LLM recognition ENABLED with model: {args.model}")
        coord_mode = "original" if args.original_coords else "compressed"
        logger.info(f"Coordinate mode: {coord_mode}")

    # Validate input file
    file = Path(args.input_file)
    if not file.exists():
        logger.error(f"Input file not found: {file}")
        exit(1)

    if file.suffix.lower() not in [".xlsx", ".xls"]:
        logger.error(f"Input file must be an Excel file (.xlsx or .xls): {file}")
        exit(1)

    logger.info(f"Input file: {file}")

    # Create output directory if it doesn't exist
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"Output directory: {output_dir}")

    # Initialize model if recognition is enabled
    model = None
    if args.recognize:
        try:
            from langchain_openai import ChatOpenAI

            # Check for custom OpenAI configuration
            openai_base_url = os.environ.get("OPENAI_BASE_URL")
            openai_api_key = os.environ.get("OPENAI_API_KEY")

            # Build ChatOpenAI initialization parameters
            model_params = {"model": args.model}

            if openai_base_url:
                model_params["base_url"] = openai_base_url
                logger.info(f"Using custom OpenAI base URL: {openai_base_url}")

            if openai_api_key:
                model_params["api_key"] = openai_api_key
                logger.info("Using API key from OPENAI_API_KEY environment variable")

            model = ChatOpenAI(**model_params)
            logger.info(f"Initialized model: {args.model}")
        except Exception as e:
            logger.error(f"Failed to initialize model: {e}")
            exit(1)

    # Process spreadsheet
    wrapper = SpreadsheetLLMWrapper()

    if wb := wrapper.read_spreadsheet(file):
        if result := wrapper.compress_spreadsheet(wb, format_aware=args.format_aware):
            # Log anchor information
            row_count = len(result.anchors.row_anchors)
            col_count = len(result.anchors.column_anchors)
            orig_rows, orig_cols = result.anchors.original_shape
            row_ratio = (row_count / orig_rows * 100) if orig_rows > 0 else 0
            col_ratio = (col_count / orig_cols * 100) if orig_cols > 0 else 0

            logger.info("=" * 60)
            logger.info("ANCHOR INFORMATION:")
            logger.info(f"  Row anchors: {row_count} (from {orig_rows} original rows)")
            logger.info(
                f"  Column anchors: {col_count} (from {orig_cols} original columns)"
            )
            logger.info(
                f"  Compression ratio: {row_ratio:.1f}% rows, {col_ratio:.1f}% columns retained"
            )
            logger.info("=" * 60)

            # Generate output file names
            base_name = (
                output_dir / file.stem
            )  # Use stem to get filename without extension
            suffix = "_format_aware" if args.format_aware else ""

            areas_file = str(base_name) + suffix + "_areas.txt"
            dict_file = str(base_name) + suffix + "_dict.txt"
            mapping_file = str(base_name) + suffix + "_mapping.json"

            # Write output files
            wrapper.write_areas(areas_file, result.areas, result.sheet_compressor)
            wrapper.write_dict(dict_file, result.compress_dict)
            wrapper.write_mapping(mapping_file, result.sheet_compressor)

            logger.info("Output files:")
            logger.info(f"  - {areas_file}")
            logger.info(f"  - {dict_file}")
            logger.info(f"  - {mapping_file}")

            # Calculate compression ratio
            original_size += os.path.getsize(file)
            new_size += os.path.getsize(areas_file)
            new_size += os.path.getsize(dict_file)
            logger.info("Compression Ratio: {:.2f}".format(original_size / new_size))

            # Run LLM recognition if enabled
            if args.recognize:
                if model is None:
                    logger.error("Model not initialized. Cannot run recognition.")
                    exit(1)

                total_anchors = len(result.anchors.row_anchors) + len(
                    result.anchors.column_anchors
                )

                logger.info("=" * 60)
                logger.info("Running LLM-based cell range recognition...")
                logger.info(
                    f"Based on anchor count ({len(result.anchors.row_anchors)} rows, {len(result.anchors.column_anchors)} columns, total: {total_anchors}), you can choose different models:"
                )
                logger.info(
                    "  - Small anchors (<50 total): Fast models like gpt-4o-mini"
                )
                logger.info(
                    "  - Medium anchors (50-200 total): Balanced models like gpt-4o"
                )
                logger.info(
                    "  - Large anchors (>200 total): Powerful models like gpt-4o or Claude"
                )
                logger.info("=" * 60)

                try:
                    if args.original_coords:
                        recognition_result = wrapper.recognize_original(
                            result.compress_dict,
                            result.sheet_compressor,
                            model,
                            args.user_prompt,
                        )
                    else:
                        recognition_result = wrapper.recognize(
                            result.compress_dict, model, args.user_prompt
                        )

                    # Display results
                    logger.info("=" * 60)
                    logger.info("RECOGNITION RESULTS:")
                    logger.info("=" * 60)
                    logger.info(f"\nReasoning:\n{recognition_result.reasoning}\n")
                    logger.info(f"Found {len(recognition_result.items)} cell ranges:")
                    for i, item in enumerate(recognition_result.items, 1):
                        title = item.title or "Untitled"
                        logger.info(f"  {i}. {title}: {item.range}")

                    # Save recognition results to file
                    recognition_file = str(base_name) + suffix + "_recognition.txt"
                    with open(recognition_file, "w", encoding="utf-8") as f:
                        f.write("=== LLM CELL RANGE RECOGNITION ===\n\n")
                        f.write(f"Model: {args.model}\n")
                        f.write(
                            f"Coordinate mode: {'original' if args.original_coords else 'compressed'}\n\n"
                        )
                        f.write(f"Reasoning:\n{recognition_result.reasoning}\n\n")
                        f.write(
                            f"Found {len(recognition_result.items)} cell ranges:\n\n"
                        )
                        for i, item in enumerate(recognition_result.items, 1):
                            title = item.title or "Untitled"
                            f.write(f"{i}. {title}: {item.range}\n")

                    logger.info(f"\nRecognition results saved to: {recognition_file}")
                    logger.info("=" * 60)

                except Exception as e:
                    logger.error(f"Recognition failed: {e}")
                    import traceback

                    logger.debug(traceback.format_exc())
        else:
            logger.error("Compression failed")
            exit(1)
    else:
        logger.error("Failed to read spreadsheet")
        exit(1)
