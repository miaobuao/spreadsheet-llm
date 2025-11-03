"""
Range compression utilities.

This module provides functions to compress specific ranges in spreadsheets
without full anchor detection, generating inverted index representations.
"""

import logging
from typing import Dict, List, Tuple

import pandas as pd

from spreadsheet_llm.index_column_converter import IndexColumnConverter
from spreadsheet_llm.sheet_compressor import SheetCompressor
from spreadsheet_llm.unified_workbook import UnifiedWorksheet

logger = logging.getLogger(__name__)


def compress_range(
    ws: UnifiedWorksheet,
    range_spec: Tuple[str, str],
    format_aware: bool = False,
) -> Dict[str, List[str]]:
    """
    Compress a specific range in a spreadsheet and return its inverted index.

    This function performs compression on a user-specified range without the
    anchor detection step. It directly encodes the given range and generates
    an inverted index mapping values/types to cell addresses.

    Supports multiple formats (.xlsx, .xlsb, .xls) through UnifiedWorksheet.

    Args:
        ws: UnifiedWorksheet instance
        range_spec: Tuple of (start_cell, end_cell) defining the range.
                   Examples: ("A1", "B10"), ("C3", "C3")
        format_aware: If True, enables format-aware aggregation.
                     Groups cells by both value AND data type.
                     If False (default), groups cells only by value.
                     Note: Format information only available for .xlsx files.

    Returns:
        Dict[str, List[str]]: Inverted index mapping values/types to cell addresses.
                             Cell addresses are absolute coordinates relative to the
                             entire sheet (e.g., "A5", "B10").

    Example:
        >>> from spreadsheet_llm.unified_workbook import create_unified_workbook
        >>> from spreadsheet_llm.range_compressor import compress_range
        >>> wb = create_unified_workbook("data.xlsx")
        >>> ws = wb.active
        >>> result = compress_range(ws, ("A1", "C10"), format_aware=True)
        >>> for key, cells in result.items():
        ...     print(f"{key}: {cells}")
        Eagles: ['A1', 'A2:A5']
        ${Integer}: ['B1:B10', 'C1:C5']
    """
    converter = IndexColumnConverter()

    logger.info(f"Compressing range on sheet: '{ws.title}'")

    # Parse range coordinates
    start_cell, end_cell = range_spec
    start_parsed = converter.parse_cell(start_cell)
    end_parsed = converter.parse_cell(end_cell)

    if not start_parsed or not end_parsed:
        raise ValueError(f"Invalid range format: {range_spec}")

    start_col, start_row = start_parsed
    end_col, end_row = end_parsed

    # Convert to 0-based indices
    start_row_idx = int(start_row) - 1
    end_row_idx = int(end_row) - 1
    start_col_idx = converter.parse_cellindex(start_col) - 1
    end_col_idx = converter.parse_cellindex(end_col) - 1

    logger.info(
        f"Range parsed: rows {start_row_idx+1}-{end_row_idx+1}, "
        f"cols {start_col_idx+1}-{end_col_idx+1} "
        f"({converter.parse_colindex(start_col_idx+1)}-{converter.parse_colindex(end_col_idx+1)})"
    )

    # Read the specified range into a DataFrame
    num_rows = end_row_idx - start_row_idx + 1
    num_cols = end_col_idx - start_col_idx + 1

    # Read the specified range directly from worksheet
    data = []
    for row_idx in range(start_row_idx, end_row_idx + 1):
        row_data = []
        for col_idx in range(start_col_idx, end_col_idx + 1):
            cell = ws.cell(row=row_idx + 1, column=col_idx + 1)  # openpyxl is 1-based
            row_data.append(cell.value)
        data.append(row_data)

    # Create DataFrame from the range data
    range_df = pd.DataFrame(data)
    range_df.columns = list(range(len(range_df.columns)))

    logger.info(f"Extracted range shape: {range_df.shape} (rows x cols)")
    logger.debug(f"Range data:\n{range_df.head()}")

    # Create SheetCompressor with manual mapping
    # Since we're not doing anchor detection, the mapping is a direct 1:1
    # from range indices to absolute sheet indices
    sheet_compressor = SheetCompressor(format_aware=format_aware)

    # Setup mapping: range index -> absolute sheet index
    sheet_compressor.row_mapping = {i: start_row_idx + i for i in range(num_rows)}
    sheet_compressor.column_mapping = {i: start_col_idx + i for i in range(num_cols)}

    logger.debug(f"Row mapping: {sheet_compressor.row_mapping}")
    logger.debug(f"Column mapping: {sheet_compressor.column_mapping}")

    # Encode the range
    markdown = sheet_compressor.encode(ws, range_df)
    logger.info(f"Encoded markdown entries: {len(markdown)}")

    # Add categories if format_aware
    if format_aware:
        markdown["Category"] = markdown["Value"].apply(
            lambda x: sheet_compressor.get_category(x)
        )
        logger.info(f"Categories assigned: {markdown['Category'].unique()}")

    # Generate inverted index
    compress_dict = sheet_compressor.inverted_index(markdown)
    logger.info(f"Inverted index created with {len(compress_dict)} unique entries")

    # Convert compressed coordinates to absolute coordinates
    # The encode() method generates addresses using compressed indices (starting from A1)
    # We need to convert these to absolute coordinates relative to the entire sheet
    absolute_dict: Dict[str, List[str]] = {}

    for key, cell_list in compress_dict.items():
        absolute_cells = []
        for cell_range in cell_list:
            # Handle both single cells and ranges
            if ":" in cell_range:
                # Range format: "A1:B5"
                start, end = cell_range.split(":")
                start_absolute = _convert_cell_to_absolute(
                    start, start_row_idx, start_col_idx, converter
                )
                end_absolute = _convert_cell_to_absolute(
                    end, start_row_idx, start_col_idx, converter
                )
                absolute_cells.append(f"{start_absolute}:{end_absolute}")
            else:
                # Single cell format: "A1"
                absolute_cell = _convert_cell_to_absolute(
                    cell_range, start_row_idx, start_col_idx, converter
                )
                absolute_cells.append(absolute_cell)

        absolute_dict[key] = absolute_cells

    logger.info(f"Converted {len(absolute_dict)} entries to absolute coordinates")
    return absolute_dict


def _convert_cell_to_absolute(
    cell: str, row_offset: int, col_offset: int, converter: IndexColumnConverter
) -> str:
    """
    Convert a cell coordinate from range-relative to sheet-absolute.

    Args:
        cell: Cell coordinate relative to range (e.g., "A1", "B5")
        row_offset: Starting row index (0-based) of the range in the sheet
        col_offset: Starting column index (0-based) of the range in the sheet
        converter: IndexColumnConverter instance

    Returns:
        Absolute cell coordinate (e.g., "D10" if range starts at D10 and cell is "A1")
    """
    parsed = converter.parse_cell(cell)
    if not parsed:
        logger.warning(f"Failed to parse cell: {cell}")
        return cell

    col_str, row_str = parsed
    # Convert to 0-based indices relative to range
    col_idx = converter.parse_cellindex(col_str) - 1
    row_idx = int(row_str) - 1

    # Add offset to get absolute indices
    abs_row_idx = row_idx + row_offset
    abs_col_idx = col_idx + col_offset

    # Convert back to cell notation (1-based)
    abs_col_str = converter.parse_colindex(abs_col_idx + 1)
    abs_row_str = str(abs_row_idx + 1)

    return f"{abs_col_str}{abs_row_str}"
