"""
Utilities for combining Excel cell addresses into ranges.

This module provides functions to intelligently combine multiple cell addresses
(e.g., ["A1", "A2", "B1", "B2"]) into compact range representations (e.g., "A1:B2").
"""

import logging

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
