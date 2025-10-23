import os
from pathlib import Path
import pandas as pd
import openpyxl
import logging
import json
from termcolor import colored
from IndexColumnConverter import IndexColumnConverter
from SheetCompressor import SheetCompressor


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


class SpreadsheetLLMWrapper:

    def __init__(self):
        return

    def read_spreadsheet(self, file):

        wb = openpyxl.load_workbook(file)
        return wb

    # Takes a file, compresses it
    def compress_spreadsheet(self, wb):
        sheet_compressor = SheetCompressor()
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

    def write_areas(self, file, areas, sheet_compressor):
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
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)

    def write_dict(self, file, dict):
        string = ""
        for key, value in dict.items():
            # value is now list[str], join with comma
            value_str = ",".join(value) if isinstance(value, list) else str(value)
            string += value_str + "," + str(key) + "|\n"
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)

    def write_mapping(self, file, sheet_compressor):
        """Write coordinate mapping from compressed to original coordinates as JSON"""
        mapping = sheet_compressor.get_coordinate_mapping()
        with open(file, "w+", encoding="utf-8") as f:
            json.dump(mapping, f, indent=2, ensure_ascii=False)


if __name__ == "__main__":

    wrapper = SpreadsheetLLMWrapper()
    file = Path(
        "/Volumes/Yang/dev/contextgen_doc/benchmark/validation/dsbench/00000011/MO14-Purple-City.xlsx"
    )
    if wb := wrapper.read_spreadsheet(file):
        if compress_result := wrapper.compress_spreadsheet(wb):
            areas, compress_dict, sheet_compressor = compress_result
            base_name = "output/" + file.name.split(".")[0]

            wrapper.write_areas(base_name + "_areas.txt", areas, sheet_compressor)
            wrapper.write_dict(base_name + "_dict.txt", compress_dict)
            wrapper.write_mapping(base_name + "_mapping.json", sheet_compressor)

            original_size += os.path.getsize(file)
            new_size += os.path.getsize(base_name + "_areas.txt")
            new_size += os.path.getsize(base_name + "_dict.txt")
            logger.info("Compression Ratio: {}".format(str(original_size / new_size)))
        else:
            logger.info("compress failed")

    else:
        logger.info("no spread sheet")
