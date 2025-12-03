"""Test helper functions for range manipulation."""

from spreadsheet_llm.cell_range_utils import get_cells_in_range


def test_get_cells_in_range_single_range():
    """Test expanding a single range."""
    result = get_cells_in_range("A1:B2")
    expected = {"A1", "A2", "B1", "B2"}
    assert result == expected


def test_get_cells_in_range_single_cell():
    """Test expanding a single cell."""
    

    result = get_cells_in_range("A1")
    expected = {"A1"}
    assert result == expected


def test_get_cells_in_range_multiple():
    """Test expanding multiple ranges."""
    

    result = get_cells_in_range("A1:A3,B5")
    expected = {"A1", "A2", "A3", "B5"}
    assert result == expected


def test_get_cells_in_range_larger():
    """Test expanding a larger range."""
    

    result = get_cells_in_range("A1:C2")
    expected = {"A1", "A2", "B1", "B2", "C1", "C2"}
    assert result == expected


def test_get_cells_in_range_column():
    """Test expanding a single column range."""
    

    result = get_cells_in_range("A1:A5")
    expected = {"A1", "A2", "A3", "A4", "A5"}
    assert result == expected


def test_get_cells_in_range_row():
    """Test expanding a single row range."""
    

    result = get_cells_in_range("A1:E1")
    expected = {"A1", "B1", "C1", "D1", "E1"}
    assert result == expected


if __name__ == "__main__":
    test_get_cells_in_range_single_range()
    test_get_cells_in_range_single_cell()
    test_get_cells_in_range_multiple()
    test_get_cells_in_range_larger()
    test_get_cells_in_range_column()
    test_get_cells_in_range_row()
    print("✓ All range helper tests passed!")
