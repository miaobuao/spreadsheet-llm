"""Test convert_compressed_to_original pure function."""

from spreadsheet_llm.cell_range_utils import convert_compressed_to_original


def test_single_cell_identity():
    """Test single cell that maps to itself."""
    mapping = {"A1": "A1", "B2": "B2"}
    result = convert_compressed_to_original("A1", mapping)
    assert result == "A1"


def test_single_cell_remapped():
    """Test single cell that maps to different coordinate."""
    mapping = {"A1": "C5", "B2": "D10"}
    result = convert_compressed_to_original("A1", mapping)
    assert result == "C5"


def test_range_identity():
    """Test range where both endpoints map to themselves."""
    mapping = {"A1": "A1", "B5": "B5"}
    result = convert_compressed_to_original("A1:B5", mapping)
    assert result == "A1:B5"


def test_range_remapped():
    """Test range where endpoints are remapped."""
    mapping = {"A1": "C10", "B5": "E20"}
    result = convert_compressed_to_original("A1:B5", mapping)
    assert result == "C10:E20"


def test_multiple_cells_comma_separated():
    """Test multiple cells separated by commas."""
    mapping = {"A1": "C10", "B2": "D15", "E5": "G25"}
    result = convert_compressed_to_original("A1,B2,E5", mapping)
    assert result == "C10,D15,G25"


def test_multiple_ranges_comma_separated():
    """Test multiple ranges separated by commas."""
    mapping = {"A1": "C10", "B5": "E20", "D3": "F15", "F8": "H30"}
    result = convert_compressed_to_original("A1:B5,D3:F8", mapping)
    assert result == "C10:E20,F15:H30"


def test_mixed_cells_and_ranges():
    """Test mix of single cells and ranges."""
    mapping = {"A1": "C10", "B2": "D15", "E5": "G25", "F8": "H30"}
    result = convert_compressed_to_original("A1,B2:E5,F8", mapping)
    assert result == "C10,D15:G25,H30"


def test_single_cell_range():
    """Test single cell expressed as range (A1:A1)."""
    mapping = {"A1": "C10"}
    result = convert_compressed_to_original("A1:A1", mapping)
    assert result == "C10:C10"


def test_cell_not_in_mapping():
    """Test cell that doesn't exist in mapping - should return as-is with warning."""
    mapping = {"A1": "C10"}
    # Cell Z99 not in mapping, should return as-is
    result = convert_compressed_to_original("Z99", mapping)
    assert result == "Z99"


def test_complex_real_world_scenario():
    """Test complex scenario from actual use case."""
    # Simulating compressed to original mapping
    mapping = {
        "A1": "A1",
        "B1": "E1",
        "C1": "J1",
        "A2": "A5",
        "B2": "E10",
        "C2": "J15",
    }

    # Multiple ranges
    result = convert_compressed_to_original("A1:C1,B2", mapping)
    assert result == "A1:J1,E10"


def test_empty_string():
    """Test empty string input."""
    mapping = {"A1": "C10"}
    result = convert_compressed_to_original("", mapping)
    assert result == ""


def test_whitespace_handling():
    """Test that whitespace is properly handled."""
    mapping = {"A1": "C10", "B2": "D15"}
    result = convert_compressed_to_original("A1 , B2", mapping)
    assert result == "C10,D15"


def test_large_coordinate_numbers():
    """Test with large row/column numbers."""
    mapping = {"AA100": "ZZ1000", "AB200": "AAA2000"}
    result = convert_compressed_to_original("AA100:AB200", mapping)
    assert result == "ZZ1000:AAA2000"


if __name__ == "__main__":
    test_single_cell_identity()
    test_single_cell_remapped()
    test_range_identity()
    test_range_remapped()
    test_multiple_cells_comma_separated()
    test_multiple_ranges_comma_separated()
    test_mixed_cells_and_ranges()
    test_single_cell_range()
    test_cell_not_in_mapping()
    test_complex_real_world_scenario()
    test_empty_string()
    test_whitespace_handling()
    test_large_coordinate_numbers()
    print("✓ All convert_compressed_to_original tests passed!")
