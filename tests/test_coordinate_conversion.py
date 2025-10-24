"""Test script for coordinate conversion functionality"""
import logging
from pathlib import Path
from spreadsheet_llm import SpreadsheetLLMWrapper

# Setup logging
logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

if __name__ == "__main__":
    # Test with the same file used in main.py
    wrapper = SpreadsheetLLMWrapper(format_aware=True)
    file = Path("/Volumes/Yang/dev/contextgen_doc/benchmark/validation/dsbench/00000011/MO14-Purple-City.xlsx")

    logger.info(f"Reading spreadsheet: {file.name}")
    wb = wrapper.read_spreadsheet(file)

    if not wb:
        logger.error("Failed to read spreadsheet")
        exit(1)

    logger.info("Compressing spreadsheet...")
    compress_result = wrapper.compress_spreadsheet(wb)

    if not compress_result:
        logger.error("Failed to compress spreadsheet")
        exit(1)

    areas, compress_dict, sheet_compressor = compress_result

    # Test coordinate conversions
    logger.info("\n" + "="*60)
    logger.info("Testing coordinate conversions:")
    logger.info("="*60)

    test_cases = [
        "A1",           # Single cell
        "B12",          # Another single cell
        "D39:D54",      # Range from dict output
        "G14:G15",      # Another range
        "B12,B39,B44",  # Multiple cells
    ]

    for compressed_coord in test_cases:
        original_coord = wrapper.convert_compressed_to_original(compressed_coord, sheet_compressor)
        logger.info(f"Compressed: {compressed_coord:20s} -> Original: {original_coord}")

    logger.info("\n" + "="*60)
    logger.info("Converting some coordinates from the dict output:")
    logger.info("="*60)

    # Show a few examples from the actual dict
    examples_from_dict = [
        ("Eagles", "B12,B39,B44,B48,B52,B54"),
        ("${Integer}", "G14:G15,H15,G16:I17"),
        ("${yyyy/mm/dd}", "D39:D54"),
    ]

    for key, compressed_coords in examples_from_dict:
        original_coords = wrapper.convert_compressed_to_original(compressed_coords, sheet_compressor)
        logger.info(f"\nKey: {key}")
        logger.info(f"  Compressed: {compressed_coords}")
        logger.info(f"  Original:   {original_coords}")
