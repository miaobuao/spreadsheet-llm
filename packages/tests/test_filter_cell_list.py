"""Test filter_cell_list_by_range function for precise filtering and global optimization."""

from spreadsheet_llm.cell_range_utils import filter_cell_list_by_range


def test_single_cell_in_range():
    """Test single cell that is in the target range."""
    result = filter_cell_list_by_range(["A1"], "A1:B2")
    assert result == ["A1"]


def test_single_cell_not_in_range():
    """Test single cell that is not in the target range."""
    result = filter_cell_list_by_range(["C3"], "A1:B2")
    assert result == []


def test_range_full_subset():
    """Test range that is fully contained in target."""
    result = filter_cell_list_by_range(["A1:A3"], "A1:B10")
    assert result == ["A1:A3"]


def test_range_partial_subset():
    """Test range with partial overlap - the key test case."""
    # Input list: ["B2:B10"]
    # Target: A1:C5
    # Expected: Only B2:B5, not the full B2:B10
    result = filter_cell_list_by_range(["B2:B10"], "A1:C5")
    assert result == ["B2:B5"], f"Expected ['B2:B5'], got {result}"


def test_range_no_overlap():
    """Test range with no overlap."""
    result = filter_cell_list_by_range(["C10:C20"], "A1:B2")
    assert result == []


def test_multiple_cells_combined():
    """Test multiple single cells get combined."""
    # Input: separate cells A1, A2, A3
    # Target: A1:A10
    # Expected: combined into A1:A3
    result = filter_cell_list_by_range(["A1", "A2", "A3"], "A1:A10")
    assert result == ["A1:A3"]


def test_global_optimization_key_feature():
    """Test GLOBAL optimization: the key improvement over old implementation."""
    # Input: ["A1:A5", "A6:A10"]  (two separate ranges)
    # Target: "A1:A10"
    # OLD behavior: would return ["A1:A5", "A6:A10"]
    # NEW behavior: returns ["A1:A10"] (globally optimized!)
    result = filter_cell_list_by_range(["A1:A5", "A6:A10"], "A1:A10")
    assert result == ["A1:A10"], f"Expected global optimization to ['A1:A10'], got {result}"


def test_mixed_cells_and_ranges():
    """Test mix of cells and ranges."""
    # Input: ["A1", "B2:B5"]
    # Target: A1:C5
    # Expected: A1 and B2:B5 (both in range)
    result = filter_cell_list_by_range(["A1", "B2:B5"], "A1:C5")
    assert set(result) == {"A1", "B2:B5"}


def test_mixed_with_partial_overlap():
    """Test mix with one range partially outside."""
    # Input: ["A1", "B2:B10"]
    # Target: A1:C5
    # Expected: A1, B2:B5 (B6:B10 filtered out)
    result = filter_cell_list_by_range(["A1", "B2:B10"], "A1:C5")
    assert set(result) == {"A1", "B2:B5"}


def test_large_range_small_intersection():
    """Test large range with small intersection."""
    # Simulates compress_dict scenario:
    # compress_dict has "${Integer}": ["B2:B100"]
    # target range is "A1:C10"
    # Expected: only B2:B10
    result = filter_cell_list_by_range(["B2:B100"], "A1:C10")
    assert result == ["B2:B10"], f"Expected ['B2:B10'], got {result}"


def test_scattered_cells_with_gap():
    """Test scattered cells with gap."""
    # Input: A1, A2, A3, A5, A6, A7 (gap at A4)
    # Target: A1:A10
    # Expected: A1:A3, A5:A7 (two ranges with gap)
    result = filter_cell_list_by_range(["A1", "A2", "A3", "A5", "A6", "A7"], "A1:A10")
    assert result == ["A1:A3", "A5:A7"]


def test_2d_range_partial():
    """Test 2D range with partial intersection."""
    # Input: ["A1:C3"]
    # Target: B1:B2
    # Expected: B1:B2
    result = filter_cell_list_by_range(["A1:C3"], "B1:B2")
    assert result == ["B1:B2"]


def test_empty_input_list():
    """Test empty input list."""
    result = filter_cell_list_by_range([], "A1:B2")
    assert result == []


def test_multiple_ranges_global_merge():
    """Test multiple ranges that should merge globally."""
    # Input: ["A1:A3", "A4:A6", "A7:A10"]
    # Target: A1:A10
    # Expected: A1:A10 (all merged)
    result = filter_cell_list_by_range(["A1:A3", "A4:A6", "A7:A10"], "A1:A10")
    assert result == ["A1:A10"]


def test_complex_scenario():
    """Test complex real-world scenario."""
    # Simulating compress_dict:
    # "${Integer}": ["A1", "A2", "B5:B10", "C1:C3"]
    # Target: A1:B6
    # Expected: A1:A2, B5:B6 (C1:C3 out of range, B7:B10 out of range)
    result = filter_cell_list_by_range(["A1", "A2", "B5:B10", "C1:C3"], "A1:B6")
    assert set(result) == {"A1:A2", "B5:B6"}


if __name__ == "__main__":
    test_single_cell_in_range()
    test_single_cell_not_in_range()
    test_range_full_subset()
    test_range_partial_subset()
    test_range_no_overlap()
    test_multiple_cells_combined()
    test_global_optimization_key_feature()
    test_mixed_cells_and_ranges()
    test_mixed_with_partial_overlap()
    test_large_range_small_intersection()
    test_scattered_cells_with_gap()
    test_2d_range_partial()
    test_empty_input_list()
    test_multiple_ranges_global_merge()
    test_complex_scenario()
    print("✓ All filter_cell_list_by_range tests passed!")
