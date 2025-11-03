"""
Utilities for combining Excel cell addresses into ranges.

This module provides functions to intelligently combine multiple cell addresses
(e.g., ["A1", "A2", "B1", "B2"]) into compact range representations (e.g., "A1:B2").
"""

import logging
import re

logger = logging.getLogger(__name__)


def col_to_index(col_str: str) -> int:
    """
    Convert Excel column letters to 0-based index.

    Examples:
        >>> col_to_index("A")
        0
        >>> col_to_index("B")
        1
        >>> col_to_index("Z")
        25
        >>> col_to_index("AA")
        26
        >>> col_to_index("AB")
        27

    Args:
        col_str: Column letters (e.g., "A", "AA", "ABC")

    Returns:
        0-based column index
    """
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


def index_to_col(index: int) -> str:
    """
    Convert 0-based column index to Excel column letters.

    Examples:
        >>> index_to_col(0)
        'A'
        >>> index_to_col(1)
        'B'
        >>> index_to_col(25)
        'Z'
        >>> index_to_col(26)
        'AA'
        >>> index_to_col(27)
        'AB'

    Args:
        index: 0-based column index

    Returns:
        Column letters (e.g., "A", "AA", "ABC")
    """
    import string

    # Modified divmod function for Excel
    def divmod_excel(n):
        a, b = divmod(n, 26)
        if b == 0:
            return a - 1, b + 26
        return a, b

    # Convert to 1-based for Excel conversion
    num = index + 1
    chars = []
    while num > 0:
        num, d = divmod_excel(num)
        chars.append(string.ascii_uppercase[d - 1])
    return "".join(reversed(chars))


def combine_cells(array: list[str]) -> list[str]:
    """
    Combine multiple Excel cell addresses into a list of ranges and individual cells.

    This function intelligently combines cell addresses:
    - Continuous rectangular regions → merged as "A1:C3"
    - Individual cells → kept as ["D5"]
    - No greedy merging - preserves exact cell references

    Examples:
        >>> combine_cells(["A1"])
        ['A1']
        >>> combine_cells(["A1", "A2", "A3"])
        ['A1:A3']
        >>> combine_cells(["A1", "A2", "B1", "B2"])
        ['A1:B2']
        >>> combine_cells(["A1", "A3", "A5"])
        ['A1', 'A3', 'A5']
        >>> combine_cells(["A1", "A2", "A3", "A5", "A7", "A8"])
        ['A1:A3', 'A5', 'A7:A8']

    Args:
        array: List of Excel cell addresses (e.g., ["A1", "B2", "C3"])

    Returns:
        List of compact string representations (ranges or individual cells)
    """
    if len(array) == 0:
        return []
    if len(array) == 1:
        return [array[0]]

    # Parse all cell addresses to get (row, col) coordinates
    coords = []
    for addr in array:
        # Parse address like "A1" -> row=0, col=0
        # Extract column letters and row number
        col_str = "".join(c for c in addr if c.isalpha())
        row_str = "".join(c for c in addr if c.isdigit())
        if not row_str or not col_str:
            continue
        row = int(row_str) - 1  # Convert to 0-based
        col = col_to_index(col_str)  # Convert to 0-based
        coords.append((row, col, addr))

    if len(coords) == 0:
        return []
    if len(coords) == 1:
        return [coords[0][2]]

    # Sort by row, then column for spatial locality
    coords.sort(key=lambda x: (x[0], x[1]))

    # Build spatial index: (row, col) -> index for O(1) lookup
    coord_map = {(row, col): i for i, (row, col, _) in enumerate(coords)}
    used = set()
    result = []

    def try_expand_rectangle(start_idx):
        """
        Try to expand a rectangle from a starting cell using intelligent exploration.
        Returns (size, min_row, min_col, max_row, max_col, indices_used).

        Algorithm: Starting from a cell, try expanding right and down to find
        the largest valid rectangle using a greedy approach with backtracking.
        """
        start_row, start_col, _ = coords[start_idx]

        # Find the maximum possible width from this starting point
        # by exploring to the right in the first row
        max_col = start_col
        while (start_row, max_col + 1) in coord_map:
            idx = coord_map[(start_row, max_col + 1)]
            if idx in used:
                break
            max_col += 1

        current_width = max_col - start_col + 1

        # Try different widths to find the optimal rectangle
        best_size = 0
        best_rect = None

        # Try different widths (from full width down to 1)
        for width in range(current_width, 0, -1):
            test_max_col = start_col + width - 1
            test_max_row = start_row

            # Find how far down we can extend with this width
            for test_row in range(start_row, start_row + 1000):
                valid_row = True
                for test_col in range(start_col, test_max_col + 1):
                    if (test_row, test_col) not in coord_map:
                        valid_row = False
                        break
                    idx = coord_map[(test_row, test_col)]
                    if idx in used:
                        valid_row = False
                        break

                if not valid_row:
                    break
                test_max_row = test_row

            # Calculate size
            height = test_max_row - start_row + 1
            size = width * height

            if size > best_size:
                best_size = size
                best_rect = (start_row, start_col, test_max_row, test_max_col)

        if best_rect and best_size > 1:
            # Collect all indices in this rectangle
            min_r, min_c, max_r, max_c = best_rect
            indices = []
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    indices.append(coord_map[(r, c)])
            return (best_size, min_r, min_c, max_r, max_c, indices)

        return (1, start_row, start_col, start_row, start_col, [start_idx])

    # Process each unused cell
    for i, (row, col, addr) in enumerate(coords):
        if i in used:
            continue

        # Try to expand from this cell
        size, min_r, min_c, max_r, max_c, indices = try_expand_rectangle(i)

        if size > 1:
            # Found a rectangle
            min_addr = f"{index_to_col(min_c)}{min_r + 1}"
            max_addr = f"{index_to_col(max_c)}{max_r + 1}"
            result.append(f"{min_addr}:{max_addr}")
            used.update(indices)
        else:
            # Single cell
            result.append(addr)
            used.add(i)

    return result


def parse_excel_range(range_str: str) -> tuple[str, int, str, int]:
    """
    Parse Excel range string into its components.

    Converts a range like "A1:D10" into separate start and end components.

    Args:
        range_str: Excel range string (e.g., "A1:D10", "B5:Z100")

    Returns:
        Tuple of (start_col, start_row, end_col, end_row)
        - start_col: Starting column letter(s) (e.g., "A", "AA")
        - start_row: Starting row number (1-indexed)
        - end_col: Ending column letter(s)
        - end_row: Ending row number (1-indexed)

    Examples:
        >>> parse_excel_range("A1:D10")
        ('A', 1, 'D', 10)
        >>> parse_excel_range("B5:Z100")
        ('B', 5, 'Z', 100)
        >>> parse_excel_range("AA1:AB50")
        ('AA', 1, 'AB', 50)

    Raises:
        ValueError: If range_str is not in valid format
    """
    if ":" not in range_str:
        raise ValueError(f"Invalid range format: '{range_str}'. Expected format like 'A1:D10'")

    start, end = range_str.split(":")

    # Parse start cell
    start_col = "".join(c for c in start if c.isalpha())
    start_row_str = "".join(c for c in start if c.isdigit())

    # Parse end cell
    end_col = "".join(c for c in end if c.isalpha())
    end_row_str = "".join(c for c in end if c.isdigit())

    # Validate
    if not start_col or not start_row_str:
        raise ValueError(f"Invalid start cell: '{start}'")
    if not end_col or not end_row_str:
        raise ValueError(f"Invalid end cell: '{end}'")

    start_row = int(start_row_str)
    end_row = int(end_row_str)

    return start_col, start_row, end_col, end_row


def box_to_range(box: tuple[int, int, int, int]) -> str:
    """
    Convert 0-indexed box coordinates to Excel range string.

    Args:
        box: Tuple of (row_start, col_start, row_end, col_end)
             All coordinates are 0-indexed

    Returns:
        Excel range string (e.g., "A1:D10")

    Examples:
        >>> box_to_range((0, 0, 9, 3))
        'A1:D10'
        >>> box_to_range((4, 1, 99, 25))
        'B5:Z100'
        >>> box_to_range((0, 26, 49, 27))
        'AA1:AB50'
        >>> box_to_range((5, 5, 5, 5))
        'F6:F6'

    Note:
        This is the inverse of parse_excel_range when combined with col_to_index.
    """
    row_start, col_start, row_end, col_end = box

    start_col = index_to_col(col_start)
    end_col = index_to_col(col_end)
    start_row = row_start + 1  # Convert to 1-indexed
    end_row = row_end + 1  # Convert to 1-indexed

    return f"{start_col}{start_row}:{end_col}{end_row}"


def parse_range_string(range_str: str) -> list[tuple[str, str]]:
    """
    Parse a range string into a list of (start_cell, end_cell) tuples.

    Args:
        range_str: Range string like "A1:B5", "A1", or "A1:B5,C3:D10"

    Returns:
        List of (start_cell, end_cell) tuples

    Examples:
        >>> parse_range_string("A1:B5")
        [('A1', 'B5')]
        >>> parse_range_string("A1")
        [('A1', 'A1')]
        >>> parse_range_string("A1:B5,C3:D10")
        [('A1', 'B5'), ('C3', 'D10')]
        >>> parse_range_string("A1,B2:B5")
        [('A1', 'A1'), ('B2', 'B5')]
    """
    result = []
    # Split by comma to handle multiple ranges
    parts = range_str.split(",")

    for part in parts:
        part = part.strip()
        if ":" in part:
            # Range format: "A1:B5"
            start, end = part.split(":", 1)
            result.append((start.strip(), end.strip()))
        else:
            # Single cell: "A1"
            result.append((part, part))

    return result


def get_cells_in_range(range_str: str) -> set[str]:
    """
    Expand a range string into a set of all individual cell addresses.

    Args:
        range_str: Range string like "A1:B5" or "A1,B2:B5"

    Returns:
        Set of cell addresses (e.g., {"A1", "A2", "B1", "B2", ...})

    Examples:
        >>> get_cells_in_range("A1:B2")
        {'A1', 'A2', 'B1', 'B2'}
        >>> get_cells_in_range("A1,B2")
        {'A1', 'B2'}
        >>> get_cells_in_range("A1:A3,B5")
        {'A1', 'A2', 'A3', 'B5'}
    """
    cells = set()
    range_tuples = parse_range_string(range_str)

    for start_cell, end_cell in range_tuples:
        # Parse start cell
        match_start = re.match(r"^([A-Z]+)(\d+)$", start_cell.upper())
        if not match_start:
            logger.warning(f"Failed to parse cell: {start_cell}")
            continue

        start_col_str = match_start.group(1)
        start_row = int(match_start.group(2))

        # Parse end cell
        match_end = re.match(r"^([A-Z]+)(\d+)$", end_cell.upper())
        if not match_end:
            logger.warning(f"Failed to parse cell: {end_cell}")
            continue

        end_col_str = match_end.group(1)
        end_row = int(match_end.group(2))

        # Convert column letters to indices
        start_col_idx = col_to_index(start_col_str)
        end_col_idx = col_to_index(end_col_str)

        # Expand range to all cells
        for col_idx in range(start_col_idx, end_col_idx + 1):
            col_str = index_to_col(col_idx)
            for row in range(start_row, end_row + 1):
                cells.add(f"{col_str}{row}")

    return cells


def convert_compressed_to_original(compressed_coord: str, mapping: dict[str, str]) -> str:
    """
    Convert compressed coordinate(s) to original coordinate(s).

    This is a pure function that uses a pre-computed mapping dictionary.
    For convenience, use SheetCompressor.convert_compressed_to_original() instead.

    Args:
        compressed_coord: Compressed coordinate string, can be:
            - Single cell: "A1", "B5"
            - Range: "A1:B5", "C3:D10"
            - Multiple ranges: "A1,B2:B5,C3"
        mapping: Dictionary mapping compressed coordinates to original coordinates
                 (e.g., {"A1": "A1", "B5": "C10", ...})

    Returns:
        Original coordinate string in the same format as input

    Examples:
        >>> mapping = {"A1": "A1", "B5": "C10"}
        >>> convert_compressed_to_original("A1", mapping)
        'A1'
        >>> convert_compressed_to_original("B5", mapping)
        'C10'
        >>> convert_compressed_to_original("A1:B5", mapping)
        'A1:C10'
        >>> convert_compressed_to_original("A1,B2:B5", mapping)
        'A1,C3:C8'
    """
    import logging

    logger = logging.getLogger(__name__)

    def convert_single_cell(cell: str) -> str:
        """Convert a single cell coordinate using the mapping table"""
        cell = cell.strip()

        # Direct lookup in mapping table
        if cell in mapping:
            return mapping[cell]
        else:
            logger.warning(f"Cell {cell} not found in mapping table")
            return cell

    def convert_range(range_str: str) -> str:
        """Convert a cell range (e.g., 'A1:B5')"""
        if ":" in range_str:
            start, end = range_str.split(":")
            return f"{convert_single_cell(start)}:{convert_single_cell(end)}"
        else:
            return convert_single_cell(range_str)

    # Handle multiple ranges separated by commas
    parts = compressed_coord.split(",")
    converted_parts = [convert_range(part.strip()) for part in parts]

    return ",".join(converted_parts)


def filter_cell_list_by_range(cell_list: list[str], target_range: str) -> list[str]:
    """
    Filter a list of cells/ranges by a target range.

    Returns globally optimized combined result.
    This function collects all cells from the list that intersect with the target range,
    then performs a single global optimization to combine them into optimal ranges.

    Args:
        cell_list: List containing cells ("A1") and/or ranges ("A1:B10")
        target_range: Target range to filter by (e.g., "A1:C5")

    Returns:
        Optimally combined list of ranges/cells.
        Empty list if no intersection.

    Examples:
        >>> filter_cell_list_by_range(["A1:A5", "A6:A10"], "A1:A10")
        ['A1:A10']  # Global optimization: merges into single range
        >>> filter_cell_list_by_range(["A1", "B2:B10"], "A1:C5")
        ['A1', 'B2:B5']
        >>> filter_cell_list_by_range(["A1", "A2", "A3"], "A1:A10")
        ['A1:A3']  # Combines consecutive cells
        >>> filter_cell_list_by_range(["B2:B10"], "C1:C10")
        []  # No intersection
    """
    target_cells = get_cells_in_range(target_range)
    intersection_cells = set()

    # Collect all intersecting cells
    for cell_or_range in cell_list:
        if ":" in cell_or_range:
            cells = get_cells_in_range(cell_or_range)
        else:
            cells = {cell_or_range}
        intersection_cells.update(cells & target_cells)

    if not intersection_cells:
        return []

    # Global optimization: combine all cells into optimal ranges
    return combine_cells(list(intersection_cells))
