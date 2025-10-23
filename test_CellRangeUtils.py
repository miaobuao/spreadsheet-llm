"""
Unit tests for CellRangeUtils module.

Run with:
    python test_CellRangeUtils.py  # Manual tests
    pytest test_CellRangeUtils.py  # Using pytest
"""

from CellRangeUtils import col_to_index, combine_cells


def test_col_to_index_single_letter():
    """Test column conversion for single letters"""
    assert col_to_index("A") == 0
    assert col_to_index("B") == 1
    assert col_to_index("Z") == 25


def test_col_to_index_double_letter():
    """Test column conversion for double letters"""
    assert col_to_index("AA") == 26
    assert col_to_index("AB") == 27
    assert col_to_index("AZ") == 51
    assert col_to_index("BA") == 52


def test_combine_cells_empty():
    """Test empty input"""
    assert combine_cells([]) == []


def test_combine_cells_single():
    """Test single cell"""
    assert combine_cells(["A1"]) == ["A1"]
    assert combine_cells(["Z99"]) == ["Z99"]


def test_combine_cells_continuous_column():
    """Test continuous cells in a single column"""
    result = combine_cells(["A1", "A2", "A3"])
    assert result == ["A1:A3"]


def test_combine_cells_continuous_row():
    """Test continuous cells in a single row"""
    result = combine_cells(["A1", "B1", "C1"])
    assert result == ["A1:C1"]


def test_combine_cells_rectangle():
    """Test rectangular regions"""
    # 2x2 rectangle
    result = combine_cells(["A1", "A2", "B1", "B2"])
    assert result == ["A1:B2"]

    # 3x4 rectangle
    cells = []
    for row in range(1, 4):  # Rows 1-3
        for col in ['A', 'B', 'C', 'D']:  # Cols A-D
            cells.append(f"{col}{row}")
    result = combine_cells(cells)
    assert result == ["A1:D3"]


def test_combine_cells_non_contiguous():
    """Test non-contiguous cells (precise, no greedy)"""
    result = combine_cells(["A1", "A3", "A5"])
    assert result == ["A1", "A3", "A5"]

    result = combine_cells(["A1", "C3", "E5"])
    assert result == ["A1", "C3", "E5"]


def test_combine_cells_mixed():
    """Test mixed continuous and non-contiguous cells"""
    # Some continuous, some not
    result = combine_cells(["A1", "A2", "A3", "A5", "A7", "A8"])
    assert result == ["A1:A3", "A5", "A7:A8"]

    # Many non-contiguous (still precise, no greedy)
    many_cells = [f"A{i}" for i in [1, 3, 5, 7, 9, 11, 13]]
    result = combine_cells(many_cells)
    assert result == ["A1", "A3", "A5", "A7", "A9", "A11", "A13"]


def test_combine_cells_unsorted_input():
    """Test that function handles unsorted input correctly"""
    # Unsorted rectangle
    result = combine_cells(["B2", "A1", "B1", "A2"])
    assert result == ["A1:B2"]

    # Unsorted column
    result = combine_cells(["A3", "A1", "A2"])
    assert result == ["A1:A3"]


def test_combine_cells_with_gaps_in_rectangle():
    """Test cells that look like a rectangle but have gaps"""
    # Missing A2 and B2 - algorithm finds two 1x2 horizontal rectangles
    cells = ["A1", "A3", "B1", "B3"]
    result = combine_cells(cells)
    # Should return two horizontal ranges (more compact than 4 individual cells)
    assert result == ["A1:B1", "A3:B3"]


def test_combine_cells_large_rectangle():
    """Test a larger rectangular region"""
    # 5x5 grid
    cells = []
    for row in range(1, 6):
        for col in ['A', 'B', 'C', 'D', 'E']:
            cells.append(f"{col}{row}")
    result = combine_cells(cells)
    assert result == ["A1:E5"]


def test_combine_cells_with_invalid_addresses():
    """Test handling of invalid cell addresses"""
    # Mix of valid and invalid (invalid should be ignored)
    result = combine_cells(["A1", "", "A2"])
    assert result == ["A1:A2"]


if __name__ == "__main__":
    # Run tests manually
    import sys

    print("Running manual tests...")
    tests = [
        test_col_to_index_single_letter,
        test_col_to_index_double_letter,
        test_combine_cells_empty,
        test_combine_cells_single,
        test_combine_cells_continuous_column,
        test_combine_cells_continuous_row,
        test_combine_cells_rectangle,
        test_combine_cells_non_contiguous,
        test_combine_cells_mixed,
        test_combine_cells_unsorted_input,
        test_combine_cells_with_gaps_in_rectangle,
        test_combine_cells_large_rectangle,
        test_combine_cells_with_invalid_addresses,
    ]

    failed = 0
    for test_func in tests:
        try:
            test_func()
            print(f"✅ {test_func.__name__}")
        except AssertionError as e:
            print(f"❌ {test_func.__name__}: {e}")
            failed += 1
        except Exception as e:
            print(f"💥 {test_func.__name__}: {e}")
            failed += 1

    print(f"\n{'='*60}")
    if failed == 0:
        print(f"✅ All {len(tests)} tests passed!")
        sys.exit(0)
    else:
        print(f"❌ {failed}/{len(tests)} tests failed")
        sys.exit(1)
