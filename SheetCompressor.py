import datetime
import numpy as np
import pandas as pd
import re
import logging
from pandas.tseries.api import guess_datetime_format  # type: ignore

from IndexColumnConverter import IndexColumnConverter

# Configure logger for this module
logger = logging.getLogger(__name__)

CATEGORIES = [
    "Integer",
    "Float",
    "Percentage",
    "Scientific Notation",
    "Date",
    "Time",
    "Currency",
    "Email",
    "Other",
]
K = 4


class SheetCompressor:
    def __init__(self):
        self.row_candidates = []
        self.column_candidates = []
        self.row_lengths = {}
        self.column_lengths = {}

    # Obtain border, fill, bold info about cell; incomplete
    def get_format(self, cell):
        format_array = []

        # Border
        if cell.border.top.style:
            format_array.append("Top Border")

        if cell.border.bottom.style:
            format_array.append("Bottom Border")

        if cell.border.left.style:
            format_array.append("Left Border")

        if cell.border.right.style:
            format_array.append("Right Border")

        # Fill
        if cell.fill.start_color and cell.fill.start_color.index != "00000000":
            format_array.append("Fill Color")

        # Bold
        if cell.font.bold:
            format_array.append("Font Bold")

        return format_array

    # Encode spreadsheet into markdown format
    def encode(self, wb, sheet):
        converter = IndexColumnConverter()
        markdown = pd.DataFrame(columns=["Address", "Value", "Format"])
        ws = wb.active  # Get the active worksheet from openpyxl workbook
        for rowindex, i in sheet.iterrows():
            for colindex, j in enumerate(sheet.columns.tolist()):
                # openpyxl uses 1-based indexing for rows and columns
                cell = ws.cell(row=rowindex + 1, column=colindex + 1)
                new_row = pd.DataFrame(
                    [
                        converter.parse_colindex(colindex + 1) + str(rowindex + 1),
                        i[j],
                        self.get_format(cell),
                    ]
                ).T
                new_row.columns = markdown.columns
                markdown = pd.concat([markdown, new_row])
        return markdown

    # Checks for identical dtypes across row/column
    def get_dtype_row(self, sheet):
        current_type = []
        for i, j in sheet.iterrows():
            if current_type != (temp := j.apply(type).to_list()):
                current_type = temp
                self.row_candidates = np.append(self.row_candidates, i)

    def get_dtype_column(self, sheet):
        current_type = []
        for i, j in enumerate(sheet.columns):
            if current_type != (temp := sheet[j].apply(type).to_list()):
                current_type = temp
                self.column_candidates = np.append(self.column_candidates, i)

    # Checks for length of text across row/column, looks for outliers, marks as candidates
    def get_length_row(self, sheet):
        for i, j in sheet.iterrows():
            self.row_lengths[i] = sum(
                j.apply(
                    lambda x: (
                        0
                        if isinstance(x, float)
                        or isinstance(x, int)
                        or isinstance(x, datetime.datetime)
                        else len(x)
                    )
                )
            )
        mean = np.mean(list(self.row_lengths.values()))
        std = np.std(list(self.row_lengths.values()))
        min = np.max(mean - 2 * std, 0)
        max = mean + 2 * std
        self.row_lengths = dict(
            (k, v) for k, v in self.row_lengths.items() if v < min or v > max
        )

    def get_length_column(self, sheet):
        for i, j in enumerate(sheet.columns):
            self.column_lengths[i] = sum(
                sheet[j].apply(
                    lambda x: (
                        0
                        if isinstance(x, float)
                        or isinstance(x, int)
                        or isinstance(x, datetime.datetime)
                        else len(x)
                    )
                )
            )
        mean = np.mean(list(self.column_lengths.values()))
        std = np.std(list(self.column_lengths.values()))
        min = np.max(mean - 2 * std, 0)
        max = mean + 2 * std
        self.column_lengths = dict(
            (k, v) for k, v in self.column_lengths.items() if v < min or v > max
        )

    def anchor(self, sheet):

        # Given num, obtain all integers from num - k to num + k inclusive
        def surrounding_k(num, k):
            return list(range(num - k, num + k + 1))

        self.get_dtype_row(sheet)
        self.get_dtype_column(sheet)
        self.get_length_row(sheet)
        self.get_length_column(sheet)

        logger.debug(f"Initial row candidates (dtype): {sorted(self.row_candidates)}")
        logger.debug(
            f"Initial column candidates (dtype): {sorted(self.column_candidates)}"
        )
        logger.debug(f"Row length outliers: {sorted(self.row_lengths.keys())}")
        logger.debug(f"Column length outliers: {sorted(self.column_lengths.keys())}")

        # Keep candidates found in EITHER dtype OR length method (use union instead of intersect)
        # This is more permissive and retains more table structure information
        self.row_candidates = np.union1d(
            list(self.row_lengths.keys()), self.row_candidates
        )
        self.column_candidates = np.union1d(
            list(self.column_lengths.keys()), self.column_candidates
        )

        logger.info(f"Row candidates after union: {sorted(self.row_candidates)}")
        logger.info(f"Column candidates after union: {sorted(self.column_candidates)}")

        # Beginning/End are candidates
        self.row_candidates = np.append(
            self.row_candidates, [0, len(sheet) - 1]
        ).astype("int32")
        self.column_candidates = np.append(
            self.column_candidates, [0, len(sheet.columns) - 1]
        ).astype("int32")

        # Get K closest rows/columns to each candidate
        self.row_candidates = np.unique(
            list(
                np.concatenate([surrounding_k(i, K) for i in self.row_candidates]).flat
            )
        )
        self.column_candidates = np.unique(
            list(
                np.concatenate(
                    [surrounding_k(i, K) for i in self.column_candidates]
                ).flat
            )
        )

        # Truncate negative/out of bounds
        self.row_candidates = self.row_candidates[
            (self.row_candidates >= 0) & (self.row_candidates < len(sheet))
        ]
        self.column_candidates = self.column_candidates[
            (self.column_candidates >= 0)
            & (self.column_candidates < len(sheet.columns))
        ]

        logger.info(
            f"Final row candidates (total {len(self.row_candidates)}): {sorted(self.row_candidates.tolist())}"
        )
        logger.info(
            f"Final column candidates (total {len(self.column_candidates)}): {sorted(self.column_candidates.tolist())}"
        )
        logger.warning(
            f"Retention rate: {len(self.row_candidates)}/{len(sheet)} rows ({len(self.row_candidates)/len(sheet)*100:.1f}%), {len(self.column_candidates)}/{len(sheet.columns)} cols ({len(self.column_candidates)/len(sheet.columns)*100:.1f}%)"
        )

        sheet = sheet.iloc[self.row_candidates, self.column_candidates]

        # Remap coordinates
        sheet = sheet.reset_index().drop(columns="index")
        sheet.columns = list(range(len(sheet.columns)))

        return sheet

    # Converts markdown to value-key pair
    def inverted_index(self, markdown):

        # Takes array of Excel cells and combines adjacent cells
        def combine_cells(array):

            # Correct version
            # 2d version of summary ranges from leetcode
            # For each row, run summary ranges to get a 1d array, then run summary ranges for each column

            # Greedy version
            if len(array) == 1:
                return array[0]
            return array[0] + ":" + array[-1]

        dictionary = {}
        for _, i in markdown.iterrows():
            if i["Value"] in dictionary:
                dictionary[i["Value"]].append(i["Address"])
            else:
                dictionary[i["Value"]] = [i["Address"]]
        dictionary = {k: v for k, v in dictionary.items() if not pd.isna(k)}
        dictionary = {k: combine_cells(v) for k, v in dictionary.items()}
        return dictionary

    # Key-Value to Value-Key for categories
    def inverted_category(self, markdown):
        dictionary = {}
        for _, i in markdown.iterrows():
            dictionary[i["Value"]] = i["Category"]
        return dictionary

    # Regex to NFS
    def get_category(self, string):
        if pd.isna(string):
            return "Other"
        if isinstance(string, float):
            return "Float"
        if isinstance(string, int):
            return "Integer"
        if isinstance(string, datetime.datetime):
            return "yyyy/mm/dd"
        if re.match(r"^-?\d+$", string):
            return "Integer"
        if re.match(r"^-?\d+\.\d+$", string):
            return "Float"
        if re.match(r"^[-+]?\d*\.?\d*%$", string) or re.match(
            r"^\d{1,3}(,\d{3})*(\.\d+)?%$", string
        ):
            return "Percentage"
        if re.match(r"^[-+]?[$]\d*\.?\d{2}$", string) or re.match(
            r"^[-+]?[$]\d{1,3}(,\d{3})*(\.\d{2})?$", string
        ):  # Michael Ash
            return "Currency"
        if re.match(r"\b-?[1-9](?:\.\d+)?[Ee][-+]?\d+\b", string):  # Michael Ash
            return "Scientific Notation"
        if re.match(
            r"^((([!#$%&'*+\-/=?^_`{|}~\w])|([!#$%&'*+\-/=?^_`{|}~\w][!#$%&'*+\-/=?^_`{|}~\.\w]{0,}[!#$%&'*+\-/=?^_`{|}~\w]))[@]\w+([-.]\w+)*\.\w+([-.]\w+)*)$",
            string,
        ):  # Dave Black RFC 2821
            return "Email"
        if datetime_format := guess_datetime_format(string):
            return datetime_format
        return "Other"

    def identical_cell_aggregation(self, sheet, dictionary):

        # Handles nan edge cases
        def replace_nan(sheet):
            if pd.isna(sheet):
                return "Other"
            else:
                return dictionary[sheet]

        # DFS for checking bounds
        def dfs(r, c, val_type):
            match = replace_nan(sheet.iloc[r, c])
            if visited[r][c] or val_type != match:
                return [r, c, r - 1, c - 1]
            visited[r][c] = True
            bounds = [r, c, r, c]
            for i in [[r - 1, c], [r, c - 1], [r + 1, c], [r, c + 1]]:
                if (
                    (i[0] < 0)
                    or (i[1] < 0)
                    or (i[0] >= len(sheet))
                    or (i[1] >= len(sheet.columns))
                ):
                    continue
                match = replace_nan(sheet.iloc[i[0], i[1]])
                if not visited[i[0]][i[1]] and val_type == match:
                    new_bounds = dfs(i[0], i[1], val_type)
                    bounds = [
                        min(new_bounds[0], bounds[0]),
                        min(new_bounds[1], bounds[1]),
                        max(new_bounds[2], bounds[2]),
                        max(new_bounds[3], bounds[3]),
                    ]
            return bounds

        m = len(sheet)
        n = len(sheet.columns)

        logger.info(f"Processing cell aggregation for sheet of size {m}x{n}")
        logger.debug(f"Unique categories in dictionary: {set(dictionary.values())}")

        visited = [[False] * n for _ in range(m)]
        areas = []

        for r in range(m):
            for c in range(n):
                if not visited[r][c]:
                    val_type = replace_nan(sheet.iloc[r, c])
                    bounds = dfs(r, c, val_type)
                    area_size = (bounds[2] - bounds[0] + 1) * (
                        bounds[3] - bounds[1] + 1
                    )
                    areas.append(
                        [(bounds[0], bounds[1]), (bounds[2], bounds[3]), val_type]
                    )
                    if len(areas) <= 20:  # Log first 20 areas
                        logger.debug(
                            f"Area {len(areas)}: rows {bounds[0]}-{bounds[2]}, cols {bounds[1]}-{bounds[3]}, type={val_type}, size={area_size}"
                        )

        logger.info(f"Total areas found: {len(areas)}")
        return areas
