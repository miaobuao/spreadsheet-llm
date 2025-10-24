"""
Utilities for combining Excel cell addresses into ranges.

This module provides functions to intelligently combine multiple cell addresses
(e.g., ["A1", "A2", "B1", "B2"]) into compact range representations (e.g., "A1:B2").
"""


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

    # Sort by row, then column
    coords.sort(key=lambda x: (x[0], x[1]))

    # Find all continuous rectangular regions using a greedy algorithm
    result = []
    used = set()

    for i, (row, col, addr) in enumerate(coords):
        if i in used:
            continue

        # Try to find the largest rectangle starting from this cell
        # Try to expand the rectangle
        best_rect = None
        best_size = 1
        best_indices = {i}

        # Try different expansions
        for end_i in range(i + 1, len(coords)):
            if end_i in used:
                continue

            # Check if we can form a rectangle from i to end_i
            temp_coords = [coords[j] for j in range(i, end_i + 1) if j not in used]

            temp_rows = sorted(set(c[0] for c in temp_coords))
            temp_cols = sorted(set(c[1] for c in temp_coords))

            # Check if consecutive
            rows_consecutive = (
                all(
                    temp_rows[k] + 1 == temp_rows[k + 1]
                    for k in range(len(temp_rows) - 1)
                )
                if len(temp_rows) > 1
                else True
            )
            cols_consecutive = (
                all(
                    temp_cols[k] + 1 == temp_cols[k + 1]
                    for k in range(len(temp_cols) - 1)
                )
                if len(temp_cols) > 1
                else True
            )

            # Check if forms a complete rectangle
            expected = len(temp_rows) * len(temp_cols)
            if rows_consecutive and cols_consecutive and len(temp_coords) == expected:
                size = len(temp_coords)
                if size > best_size:
                    best_size = size
                    best_rect = (temp_coords[0][2], temp_coords[-1][2])
                    best_indices = set(range(i, end_i + 1)) - used

        if best_rect and best_size > 1:
            # Found a rectangle
            result.append(f"{best_rect[0]}:{best_rect[1]}")
            used.update(best_indices)
        else:
            # Single cell
            result.append(addr)
            used.add(i)

    return result


if __name__ == "__main__":
    # Simple manual tests
    print("Running manual tests...")

    # Test 1: Single cell
    result = combine_cells(["A1"])
    print(f"Test 1 - Single cell: {result}")
    assert result == ["A1"], f"Expected ['A1'], got '{result}'"

    # Test 2: Continuous column
    result = combine_cells(["A1", "A2", "A3"])
    print(f"Test 2 - Continuous column: {result}")
    assert result == ["A1:A3"], f"Expected ['A1:A3'], got '{result}'"

    # Test 3: Rectangular region
    result = combine_cells(["A1", "A2", "B1", "B2"])
    print(f"Test 3 - Rectangular region: {result}")
    assert result == ["A1:B2"], f"Expected ['A1:B2'], got '{result}'"

    # Test 4: Non-contiguous cells (precise, no greedy)
    result = combine_cells(["A1", "A3", "A5"])
    print(f"Test 4 - Non-contiguous: {result}")
    assert result == ["A1", "A3", "A5"], f"Expected ['A1', 'A3', 'A5'], got '{result}'"

    # Test 5: Mixed continuous and non-contiguous
    result = combine_cells(["A1", "A2", "A3", "A5", "A7", "A8"])
    print(f"Test 5 - Mixed: {result}")
    assert result == [
        "A1:A3",
        "A5",
        "A7:A8",
    ], f"Expected ['A1:A3', 'A5', 'A7:A8'], got '{result}'"

    # Test 6: Larger rectangular region
    cells = []
    for row in range(1, 4):  # Rows 1-3
        for col in ["A", "B", "C", "D"]:  # Cols A-D
            cells.append(f"{col}{row}")
    result = combine_cells(cells)
    print(f"Test 6 - Large rectangle (3x4): {result}")
    assert result == ["A1:D3"], f"Expected ['A1:D3'], got '{result}'"

    # Test 7: Column index conversion
    assert col_to_index("A") == 0, "A should be 0"
    assert col_to_index("Z") == 25, "Z should be 25"
    assert col_to_index("AA") == 26, "AA should be 26"
    assert col_to_index("AB") == 27, "AB should be 27"
    print("Test 7 - Column conversion: PASSED")

    print("\n✅ All tests passed!")
