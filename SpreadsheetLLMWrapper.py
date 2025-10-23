import json
import logging
import pandas as pd
import openpyxl
from IndexColumnConverter import IndexColumnConverter
from SheetCompressor import SheetCompressor

logger = logging.getLogger(__name__)


class SpreadsheetLLMWrapper:

    def __init__(self, format_aware: bool = False):
        """
        Initialize SpreadsheetLLMWrapper.

        Args:
            format_aware: If True, enables format-aware aggregation in dict output.
                         Groups cells by both value AND data type (e.g., "100 (Integer)").
                         If False (default), groups cells only by value.
        """
        self.format_aware = format_aware

    def read_spreadsheet(self, file):

        wb = openpyxl.load_workbook(file)
        return wb

    # Takes a file, compresses it
    def compress_spreadsheet(self, wb):
        sheet_compressor = SheetCompressor(format_aware=self.format_aware)
        # header=None: treat all rows as data, don't auto-generate column names
        sheet = pd.read_excel(wb, engine="openpyxl", header=None)
        # sheet = sheet.apply(
        #     lambda x: x.str.replace("\n", "<br>") if x.dtype == "object" else x
        # )

        # Reset index and column names to integers
        sheet = sheet.reset_index(drop=True)
        sheet.columns = list(range(len(sheet.columns)))

        logger.info(f"Original sheet shape: {sheet.shape} (rows x cols)")
        logger.debug(f"First 5 rows:\n{sheet.head()}")

        # Structural-anchor-based Extraction
        sheet = sheet_compressor.anchor(sheet)
        logger.info(f"After anchor, sheet shape: {sheet.shape} (rows x cols)")

        # Encoding
        markdown = sheet_compressor.encode(
            wb, sheet
        )  # Paper encodes first then anchors; I chose to do this in reverse
        logger.info(f"Encoded markdown shape: {markdown.shape}")
        logger.debug(f"Markdown columns: {markdown.columns.tolist()}")
        logger.debug(f"First 10 markdown entries:\n{markdown.head(10)}")

        # Data-Format Aggregation
        markdown["Category"] = markdown["Value"].apply(
            lambda x: sheet_compressor.get_category(x)
        )
        category_dict = sheet_compressor.inverted_category(markdown)
        logger.info(f"Categories found: {set(category_dict.values())}")
        logger.info(f"Number of unique values: {len(category_dict)}")

        try:
            areas = sheet_compressor.identical_cell_aggregation(sheet, category_dict)
            logger.info(f"Number of areas identified: {len(areas)}")
            logger.debug("First 10 areas:")
            for i, area in enumerate(areas[:10]):
                logger.debug(f"  Area {i}: {area}")
        except RecursionError:
            logger.error("RecursionError in identical_cell_aggregation")
            return

        # Inverted-index Translation
        compress_dict = sheet_compressor.inverted_index(markdown)
        logger.info(f"Compress dict entries: {len(compress_dict)}")

        return areas, compress_dict, sheet_compressor

    def serialize_areas(self, areas, sheet_compressor):
        """Serialize areas to string representation.

        Args:
            areas: List of area tuples
            sheet_compressor: SheetCompressor instance with row/column mappings

        Returns:
            Serialized string representation of areas
        """
        string = ""
        converter = IndexColumnConverter()
        for i in areas:
            # Map compressed indices back to original indices
            original_row_start = sheet_compressor.row_mapping[i[0][0]]
            original_col_start = sheet_compressor.column_mapping[i[0][1]]
            original_row_end = sheet_compressor.row_mapping[i[1][0]]
            original_col_end = sheet_compressor.column_mapping[i[1][1]]

            string += (
                "("
                + i[2]
                + "|"
                + converter.parse_colindex(original_col_start + 1)
                + str(original_row_start + 1)
                + ":"
                + converter.parse_colindex(original_col_end + 1)
                + str(original_row_end + 1)
                + "), "
            )
        return string

    def write_areas(self, file, areas, sheet_compressor):
        """Write serialized areas to file."""
        string = self.serialize_areas(areas, sheet_compressor)
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)

    def serialize_dict(self, dict):
        """Serialize dictionary to string representation.

        Args:
            dict: Dictionary to serialize

        Returns:
            Serialized string representation of dictionary
        """
        string = ""
        for key, value in dict.items():
            # Skip empty keys
            if not key or str(key).strip() == "":
                continue
            # value is now list[str], join with comma
            value_str = ",".join(value) if isinstance(value, list) else str(value)
            string += str(key) + "|" + value_str + "\n"
        return string

    def write_dict(self, file, dict):
        """Write serialized dictionary to file."""
        string = self.serialize_dict(dict)
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)

    def serialize_mapping(self, sheet_compressor):
        """Serialize coordinate mapping to JSON string.

        Args:
            sheet_compressor: SheetCompressor instance with row/column mappings

        Returns:
            JSON string representation of coordinate mapping
        """
        mapping = sheet_compressor.get_coordinate_mapping()
        return json.dumps(mapping, indent=2, ensure_ascii=False)

    def write_mapping(self, file, sheet_compressor):
        """Write coordinate mapping from compressed to original coordinates as JSON"""
        mapping_str = self.serialize_mapping(sheet_compressor)
        with open(file, "w+", encoding="utf-8") as f:
            f.write(mapping_str)

    def convert_compressed_to_original(self, compressed_coord: str, sheet_compressor) -> str:
        """
        Convert compressed coordinate(s) to original coordinate(s).

        Args:
            compressed_coord: Compressed coordinate string, can be:
                - Single cell: "A1", "B5"
                - Range: "A1:B5", "C3:D10"
                - Multiple ranges: "A1,B2:B5,C3"
            sheet_compressor: SheetCompressor instance with row/column mappings

        Returns:
            Original coordinate string in the same format as input

        Examples:
            "A1" -> "A1" (if A1 maps to A1)
            "B5" -> "C10" (if compressed B5 maps to original C10)
            "A1:B3" -> "A1:C5" (range conversion)
            "A1,B2:B5" -> "A1,C3:C8" (multiple ranges)
        """
        converter = IndexColumnConverter()

        def convert_single_cell(cell: str) -> str:
            """Convert a single cell coordinate"""
            # Parse cell coordinate (e.g., "A1" -> col=0, row=0)
            match = converter.parse_cell(cell)
            if not match:
                logger.warning(f"Invalid cell format: {cell}")
                return cell

            col_str, row_str = match
            compressed_col = converter.parse_cellindex(col_str) - 1  # Convert to 0-based
            compressed_row = int(row_str) - 1  # Convert to 0-based

            # Map to original indices
            original_row = sheet_compressor.row_mapping.get(compressed_row, compressed_row)
            original_col = sheet_compressor.column_mapping.get(compressed_col, compressed_col)

            # Convert back to cell notation
            original_col_str = converter.parse_colindex(original_col + 1)
            original_row_str = str(original_row + 1)

            return f"{original_col_str}{original_row_str}"

        def convert_range(range_str: str) -> str:
            """Convert a cell range (e.g., 'A1:B5')"""
            if ':' in range_str:
                start, end = range_str.split(':')
                return f"{convert_single_cell(start)}:{convert_single_cell(end)}"
            else:
                return convert_single_cell(range_str)

        # Handle multiple ranges separated by commas
        parts = compressed_coord.split(',')
        converted_parts = [convert_range(part.strip()) for part in parts]

        return ','.join(converted_parts)
