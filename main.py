import os
from pathlib import Path
import pandas as pd
import openpyxl
import logging

from IndexColumnConverter import IndexColumnConverter
from SheetCompressor import SheetCompressor

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)
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

        sheet.to_excel("./compressed.xlsx")

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

        return areas, compress_dict

    def write_areas(self, file, areas):
        string = ""
        converter = IndexColumnConverter()
        for i in areas:
            string += (
                "("
                + i[2]
                + "|"
                + converter.parse_colindex(i[0][1] + 1)
                + str(i[0][0] + 1)
                + ":"
                + converter.parse_colindex(i[1][1] + 1)
                + str(i[1][0] + 1)
                + "), "
            )
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)

    def write_dict(self, file, dict):
        string = ""
        for key, value in dict.items():
            string += str(value) + "," + str(key) + "|"
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)


if __name__ == "__main__":

    wrapper = SpreadsheetLLMWrapper()
    file = Path(
        "/Volumes/Yang/dev/contextgen_doc/benchmark/validation/dsbench/00000011/MO14-Purple-City.xlsx"
    )
    if wb := wrapper.read_spreadsheet(file):
        if compress_result := wrapper.compress_spreadsheet(wb):
            areas, compress_dict = compress_result
            print(compress_dict)

            wrapper.write_areas(
                "output/" + file.name.split(".")[0] + "_areas.txt", areas
            )
            wrapper.write_dict(
                "output/" + file.name.split(".")[0] + "_dict.txt", compress_dict
            )
            original_size += os.path.getsize(file)
            new_size += os.path.getsize(
                "output/" + file.name.split(".")[0] + "_areas.txt"
            )
            new_size += os.path.getsize(
                "output/" + file.name.split(".")[0] + "_dict.txt"
            )
            print("Compression Ratio: {}".format(str(original_size / new_size)))
        else:
            print("compress failed")

    else:
        print("no spread sheet")
