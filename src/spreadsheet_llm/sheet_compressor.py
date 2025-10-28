import datetime
import logging
import re
from typing import Any

import numpy as np
import pandas as pd
from pandas.tseries.api import guess_datetime_format  # type: ignore

from spreadsheet_llm.cell_range_utils import combine_cells
from spreadsheet_llm.index_column_converter import IndexColumnConverter

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
    def __init__(self, format_aware: bool = False):
        """
        Initialize SheetCompressor.

        Args:
            format_aware: If True, inverted_index will use format-aware aggregation,
                         grouping cells by both value AND data type (category).
                         If False (default), groups cells only by value.
        """
        self.format_aware = format_aware
        # These will be populated by the anchor() method
        self.row_candidates = np.array([])
        self.column_candidates = np.array([])
        self.row_lengths = {}
        self.column_lengths = {}
        self.row_non_null_candidates = np.array([])
        self.column_non_null_candidates = np.array([])
        # Mapping tables populated after compression
        self.row_mapping = {}  # Maps compressed row index -> original row index
        self.column_mapping = (
            {}
        )  # Maps compressed column index -> original column index

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
                # Map compressed indices back to original indices
                original_row = self.row_mapping[rowindex]
                original_col = self.column_mapping[colindex]

                # openpyxl uses 1-based indexing for rows and columns
                # Use original indices to get cell format from workbook
                cell = ws.cell(row=original_row + 1, column=original_col + 1)

                # Generate address using COMPRESSED indices (not original)
                # This keeps the dict compact and avoids sparse coordinates
                address = converter.parse_colindex(colindex + 1) + str(rowindex + 1)

                new_row = pd.DataFrame(
                    [
                        address,
                        i[j],
                        self.get_format(cell),
                    ]
                ).T
                new_row.columns = markdown.columns
                markdown = pd.concat([markdown, new_row])
        return markdown

    # Checks for identical dtypes across row/column
    def get_dtype_row(self, sheet):
        """Detect rows where data type pattern changes and update row_candidates."""
        current_type = []
        for i, j in sheet.iterrows():
            if current_type != (temp := j.apply(type).to_list()):
                current_type = temp
                self.row_candidates = np.append(self.row_candidates, i)

    def get_dtype_column(self, sheet):
        """Detect columns where data type pattern changes and update column_candidates."""
        current_type = []
        for i, j in enumerate(sheet.columns):
            if current_type != (temp := sheet[j].apply(type).to_list()):
                current_type = temp
                self.column_candidates = np.append(self.column_candidates, i)

    # Checks for non-null cell count changes across rows
    def get_non_null_count_row(self, sheet):
        """Detect rows where the number of non-null cells changes and update row_non_null_candidates."""
        prev_count = None

        for i, row in sheet.iterrows():
            count = row.notna().sum()  # Count non-NaN cells
            if prev_count is not None and count != prev_count:
                self.row_non_null_candidates = np.append(
                    self.row_non_null_candidates, i
                )
                logger.debug(
                    f"Row {i}: non-null count changed from {prev_count} to {count}"
                )
            prev_count = count

    # Checks for non-null cell count changes across columns
    def get_non_null_count_column(self, sheet):
        """Detect columns where the number of non-null cells changes and update column_non_null_candidates."""
        prev_count = None

        for col_idx in sheet.columns:
            count = sheet[col_idx].notna().sum()
            if prev_count is not None and count != prev_count:
                self.column_non_null_candidates = np.append(
                    self.column_non_null_candidates, col_idx
                )
                logger.debug(
                    f"Column {col_idx}: non-null count changed from {prev_count} to {count}"
                )
            prev_count = count

    # Checks for length of text across row/column, looks for outliers, marks as candidates
    def get_length_row(self, sheet):
        """Detect rows with outlier text lengths (mean ± 2*std) and update row_lengths."""
        for i, j in sheet.iterrows():
            self.row_lengths[i] = sum(
                j.apply(
                    lambda x: (
                        0
                        if isinstance(x, float)
                        or isinstance(x, int)
                        or isinstance(x, datetime.datetime)
                        or isinstance(x, datetime.time)
                        else len(x)
                    )
                )
            )
        mean = np.mean(list(self.row_lengths.values()))
        std = np.std(list(self.row_lengths.values()))
        min_threshold = max(mean - 2 * std, 0)
        max_threshold = mean + 2 * std
        self.row_lengths = dict(
            (k, v)
            for k, v in self.row_lengths.items()
            if v < min_threshold or v > max_threshold
        )

    def get_length_column(self, sheet):
        """Detect columns with outlier text lengths (mean ± 2*std) and update column_lengths."""
        for i, j in enumerate(sheet.columns):
            self.column_lengths[i] = sum(
                sheet[j].apply(
                    lambda x: (
                        0
                        if isinstance(x, float)
                        or isinstance(x, int)
                        or isinstance(x, datetime.datetime)
                        or isinstance(x, datetime.time)
                        else len(x)
                    )
                )
            )
        mean = np.mean(list(self.column_lengths.values()))
        std = np.std(list(self.column_lengths.values()))
        min_threshold = max(mean - 2 * std, 0)
        max_threshold = mean + 2 * std
        self.column_lengths = dict(
            (k, v)
            for k, v in self.column_lengths.items()
            if v < min_threshold or v > max_threshold
        )

    def anchor(self, sheet):

        # Given num, obtain all integers from num - k to num + k inclusive
        def surrounding_k(num, k):
            return list(range(num - k, num + k + 1))

        # Run all detection methods (each method updates instance attributes)
        self.get_dtype_row(sheet)
        self.get_dtype_column(sheet)
        self.get_length_row(sheet)
        self.get_length_column(sheet)
        self.get_non_null_count_row(sheet)
        self.get_non_null_count_column(sheet)

        logger.debug(
            f"Initial row candidates (dtype): {sorted(self.row_candidates.tolist()) if len(self.row_candidates) > 0 else []}"
        )
        logger.debug(
            f"Initial column candidates (dtype): {sorted(self.column_candidates.tolist()) if len(self.column_candidates) > 0 else []}"
        )
        logger.debug(f"Row length outliers: {sorted(self.row_lengths.keys())}")
        logger.debug(f"Column length outliers: {sorted(self.column_lengths.keys())}")
        logger.debug(
            f"Row non-null count changes: {sorted(self.row_non_null_candidates.tolist()) if len(self.row_non_null_candidates) > 0 else []}"
        )
        logger.debug(
            f"Column non-null count changes: {sorted(self.column_non_null_candidates.tolist()) if len(self.column_non_null_candidates) > 0 else []}"
        )

        # Keep candidates found in ANY method (dtype, length, or non-null count)
        # This is more permissive and retains more table structure information
        self.row_candidates = np.union1d(
            np.union1d(list(self.row_lengths.keys()), self.row_candidates),
            self.row_non_null_candidates,
        )
        self.column_candidates = np.union1d(
            np.union1d(list(self.column_lengths.keys()), self.column_candidates),
            self.column_non_null_candidates,
        )

        logger.info(
            f"Row candidates after union (dtype + length + non-null): {sorted(self.row_candidates)}"
        )
        logger.info(
            f"Column candidates after union (dtype + length + non-null): {sorted(self.column_candidates)}"
        )

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
        # Calculate retention rates with division by zero protection
        row_retention_pct = (
            len(self.row_candidates) / len(sheet) * 100 if len(sheet) > 0 else 0
        )
        col_retention_pct = (
            len(self.column_candidates) / len(sheet.columns) * 100
            if len(sheet.columns) > 0
            else 0
        )
        logger.warning(
            f"Retention rate: {len(self.row_candidates)}/{len(sheet)} rows ({row_retention_pct:.1f}%), "
            f"{len(self.column_candidates)}/{len(sheet.columns)} cols ({col_retention_pct:.1f}%)"
        )

        # Save the original row and column indices before remapping
        original_row_indices = self.row_candidates.tolist()
        original_col_indices = self.column_candidates.tolist()

        sheet = sheet.iloc[self.row_candidates, self.column_candidates]

        # Create mapping: compressed index -> original index
        self.row_mapping = {
            i: original_row_indices[i] for i in range(len(original_row_indices))
        }
        self.column_mapping = {
            i: original_col_indices[i] for i in range(len(original_col_indices))
        }

        logger.debug(f"Row mapping (compressed->original): {self.row_mapping}")
        logger.debug(f"Column mapping (compressed->original): {self.column_mapping}")

        # Remap coordinates
        sheet = sheet.reset_index().drop(columns="index")
        sheet.columns = list(range(len(sheet.columns)))

        return sheet

    # Converts markdown to value-key pair
    def inverted_index(self, markdown: pd.DataFrame):
        """
        Create inverted index from markdown.

        Args:
            markdown: DataFrame with columns Address, Value, and optionally Category

        Returns:
            Dictionary mapping values or categories to cell addresses

        Behavior:
            - If format_aware=False: Groups cells by value only
              Example: {"Eagles": ["A1:A3", "B5"]}

            - If format_aware=True: Smart aggregation
              * "Other" type: Groups by value (preserve descriptive content)
                Example: {"Eagles": ["A1:A3"], "Teams": ["B1"]}
              * Data types: Groups by category (compress numbers/dates)
                Example: {"Integer": ["C1:C10"], "yyyy/mm/dd": ["D1:D5"]}
        """
        dictionary: dict[Any, list[str]] = {}

        # Check if we should use format-aware aggregation
        use_category = self.format_aware and "Category" in markdown.columns

        if use_category:
            logger.info(
                "Using format-aware aggregation (Other uses values, data types use categories)"
            )
        else:
            logger.info("Using simple aggregation (grouping by value only)")

        for _, row in markdown.iterrows():
            value = row["Value"]

            # Skip NaN values
            if pd.isna(value):
                continue

            # Create key based on mode
            if use_category:
                category = row["Category"]
                # For "Other" type, use actual value (preserve descriptive content)
                # For data types (Integer, Float, Date, etc.), use category for compression
                if category == "Other":
                    key = value
                else:
                    # Wrap type in ${} to distinguish from literal values
                    key = f"${{{category}}}"
            else:
                # Use value only for simple mode
                key = value

            # Add address to dictionary
            if key in dictionary:
                dictionary[key].append(row["Address"])
            else:
                dictionary[key] = [row["Address"]]

        # Combine cells and format output
        res: dict[Any, list[str]] = {}
        total_keys = len(dictionary)
        for idx, (k, v) in enumerate(dictionary.items(), 1):
            # Log progress for keys with many cells or periodically
            cell_count = len(v)
            if cell_count > 100:
                logger.info(f"Processing key {idx}/{total_keys}: '{k}' with {cell_count} cells...")
            elif idx % 50 == 0:
                logger.debug(f"Progress: {idx}/{total_keys} keys processed")

            combined = combine_cells(v)  # Returns list[str]
            res[k] = combined

            # Log first 5 examples to show how cells are combined
            if len(res) <= 5:
                logger.debug(f"Key '{k}': {len(v)} cells -> {combined}")

        logger.info(f"Inverted index created with {len(res)} unique entries")
        return res

    # Get coordinate mapping from compressed to original
    def get_coordinate_mapping(self):
        """
        Returns the mapping from compressed coordinates to original coordinates.
        Format: {"A1": "A1", "E6": "Q31", ...}
        Each compressed cell coordinate maps directly to its original cell coordinate.
        """
        converter = IndexColumnConverter()
        mapping = {}

        # Generate mapping for all compressed cells
        for compressed_row_idx, original_row_idx in self.row_mapping.items():
            for compressed_col_idx, original_col_idx in self.column_mapping.items():
                # Convert indices to cell notation (e.g., "A1", "B5")
                compressed_col = converter.parse_colindex(compressed_col_idx + 1)
                compressed_row = str(compressed_row_idx + 1)
                compressed_cell = f"{compressed_col}{compressed_row}"

                original_col = converter.parse_colindex(original_col_idx + 1)
                original_row = str(original_row_idx + 1)
                original_cell = f"{original_col}{original_row}"

                mapping[compressed_cell] = original_cell

        logger.info(
            f"Coordinate mapping: {len(mapping)} cells ({len(self.row_mapping)} rows × {len(self.column_mapping)} columns)"
        )
        return mapping

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
        if isinstance(string, datetime.time):
            return "Time"

        # Convert to string for regex matching
        string = str(string)

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

        # Iterative DFS using stack (avoids RecursionError for large areas)
        def dfs_iterative(start_r, start_c, val_type):
            """Non-recursive DFS to find bounds of connected cells with same type."""
            stack = [(start_r, start_c)]
            bounds = [
                start_r,
                start_c,
                start_r,
                start_c,
            ]  # [min_r, min_c, max_r, max_c]

            while stack:
                r, c = stack.pop()

                # Skip if already visited or out of bounds
                if (
                    r < 0
                    or c < 0
                    or r >= len(sheet)
                    or c >= len(sheet.columns)
                    or visited[r][c]
                ):
                    continue

                # Check if cell type matches
                match = replace_nan(sheet.iloc[r, c])
                if val_type != match:
                    continue

                # Mark as visited and update bounds
                visited[r][c] = True
                bounds[0] = min(bounds[0], r)  # min_r
                bounds[1] = min(bounds[1], c)  # min_c
                bounds[2] = max(bounds[2], r)  # max_r
                bounds[3] = max(bounds[3], c)  # max_c

                # Add neighbors to stack (up, down, left, right)
                stack.append((r - 1, c))
                stack.append((r + 1, c))
                stack.append((r, c - 1))
                stack.append((r, c + 1))

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
                    bounds = dfs_iterative(r, c, val_type)
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
