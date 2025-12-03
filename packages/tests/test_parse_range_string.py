"""Test parse_range_string helper function."""

from spreadsheet_llm.cell_range_utils import parse_range_string


def test_parse_range_string_single_range():
    """Test parsing a single range."""
    result = parse_range_string("A1:B5")
    assert result == [("A1", "B5")]


def test_parse_range_string_single_cell():
    """Test parsing a single cell."""
    result = parse_range_string("A1")
    assert result == [("A1", "A1")]


def test_parse_range_string_multiple_ranges():
    """Test parsing multiple ranges."""
    result = parse_range_string("A1:B5,C3:D10")
    assert result == [("A1", "B5"), ("C3", "D10")]


def test_parse_range_string_mixed():
    """Test parsing mixed ranges and cells."""
    result = parse_range_string("A1,B2:B5,C3")
    assert result == [("A1", "A1"), ("B2", "B5"), ("C3", "C3")]


def test_parse_range_string_with_spaces():
    """Test parsing ranges with spaces."""
    result = parse_range_string("A1:B5 , C3:D10")
    assert result == [("A1", "B5"), ("C3", "D10")]


def test_parse_range_string_complex():
    """Test parsing complex range strings."""
    result = parse_range_string("A1:A3,B5,C7:D10,E1")
    assert result == [("A1", "A3"), ("B5", "B5"), ("C7", "D10"), ("E1", "E1")]


if __name__ == "__main__":
    test_parse_range_string_single_range()
    test_parse_range_string_single_cell()
    test_parse_range_string_multiple_ranges()
    test_parse_range_string_mixed()
    test_parse_range_string_with_spaces()
    test_parse_range_string_complex()
    print("✓ All parse_range_string tests passed!")
