"""
Comprehensive test suite for cell_range_utils module.

Tests cover:
- col_to_index: Convert Excel column letters to 0-based index
- index_to_col: Convert 0-based index to Excel column letters
- combine_cells: Combine cell addresses into compact ranges
  - Basic functionality
  - Edge cases
  - Precision requirements
  - Performance benchmarks
  - Complex real-world scenarios
"""

import sys
import time
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from spreadsheet_llm.cell_range_utils import col_to_index, index_to_col, combine_cells


class TestColToIndex:
    """Test col_to_index function: Excel column letters → 0-based index."""

    def test_single_letter_columns(self):
        """Test A-Z conversion."""
        assert col_to_index("A") == 0, "A should be 0"
        assert col_to_index("B") == 1, "B should be 1"
        assert col_to_index("Z") == 25, "Z should be 25"

    def test_double_letter_columns(self):
        """Test AA-AZ conversion."""
        assert col_to_index("AA") == 26, "AA should be 26"
        assert col_to_index("AB") == 27, "AB should be 27"
        assert col_to_index("AZ") == 51, "AZ should be 51"

    def test_triple_letter_columns(self):
        """Test AAA-ZZZ conversion."""
        assert col_to_index("BA") == 52, "BA should be 52"
        assert col_to_index("AAA") == 702, "AAA should be 702"
        assert col_to_index("ZZZ") == 18277, "ZZZ should be 18277"

    def test_case_insensitivity(self):
        """Test lowercase input."""
        assert col_to_index("a") == 0, "a should be 0"
        assert col_to_index("aa") == 26, "aa should be 26"
        assert col_to_index("Az") == 51, "Az should be 51"

    def test_common_columns(self):
        """Test commonly used columns."""
        assert col_to_index("C") == 2, "C should be 2"
        assert col_to_index("D") == 3, "D should be 3"
        assert col_to_index("J") == 9, "J should be 9"
        assert col_to_index("AC") == 28, "AC should be 28"


class TestIndexToCol:
    """Test index_to_col function: 0-based index → Excel column letters."""

    def test_single_letter_columns(self):
        """Test 0-25 → A-Z conversion."""
        assert index_to_col(0) == "A", "0 should be A"
        assert index_to_col(1) == "B", "1 should be B"
        assert index_to_col(25) == "Z", "25 should be Z"

    def test_double_letter_columns(self):
        """Test 26-51 → AA-AZ conversion."""
        assert index_to_col(26) == "AA", "26 should be AA"
        assert index_to_col(27) == "AB", "27 should be AB"
        assert index_to_col(51) == "AZ", "51 should be AZ"

    def test_triple_letter_columns(self):
        """Test larger indices."""
        assert index_to_col(52) == "BA", "52 should be BA"
        assert index_to_col(702) == "AAA", "702 should be AAA"
        assert index_to_col(18277) == "ZZZ", "18277 should be ZZZ"

    def test_common_columns(self):
        """Test commonly used column indices."""
        assert index_to_col(2) == "C", "2 should be C"
        assert index_to_col(3) == "D", "3 should be D"
        assert index_to_col(9) == "J", "9 should be J"
        assert index_to_col(28) == "AC", "28 should be AC"

    def test_round_trip_conversion(self):
        """Test that col_to_index and index_to_col are inverse functions."""
        for i in range(0, 1000, 7):  # Test 0, 7, 14, ..., 994
            col = index_to_col(i)
            assert col_to_index(col) == i, f"Round trip failed for {i} → {col} → {col_to_index(col)}"

    def test_boundary_cases(self):
        """Test boundary cases."""
        assert index_to_col(0) == "A", "First column should be A"
        assert index_to_col(25) == "Z", "Last single-letter column should be Z"
        assert index_to_col(26) == "AA", "First double-letter column should be AA"


class TestBasicFunctionality:
    """Test basic cell combination scenarios."""

    def test_empty_array(self):
        """Empty input should return empty output."""
        result = combine_cells([])
        assert result == [], f"Expected [], got {result}"

    def test_single_cell(self):
        """Single cell should return itself."""
        result = combine_cells(["A1"])
        assert result == ["A1"], f"Expected ['A1'], got {result}"

    def test_continuous_column(self):
        """Continuous column cells should form a range."""
        result = combine_cells(["A1", "A2", "A3", "A4", "A5"])
        assert result == ["A1:A5"], f"Expected ['A1:A5'], got {result}"

    def test_continuous_row(self):
        """Continuous row cells should form a range."""
        result = combine_cells(["A1", "B1", "C1", "D1"])
        assert result == ["A1:D1"], f"Expected ['A1:D1'], got {result}"

    def test_2x2_block(self):
        """2x2 block should form a single range."""
        result = combine_cells(["A1", "B1", "A2", "B2"])
        assert result == ["A1:B2"], f"Expected ['A1:B2'], got {result}"

    def test_3x3_block(self):
        """3x3 block should form a single range."""
        cells = ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"]
        result = combine_cells(cells)
        assert result == ["A1:C3"], f"Expected ['A1:C3'], got {result}"

    def test_unordered_input(self):
        """Algorithm should handle unordered input correctly."""
        result = combine_cells(["B2", "A1", "B1", "A2"])
        assert result == ["A1:B2"], f"Expected ['A1:B2'], got {result}"


class TestEdgeCases:
    """Test edge cases and boundary conditions."""

    def test_single_row_multiple_ranges(self):
        """Non-continuous cells in same row should form multiple ranges."""
        result = combine_cells(["A1", "B1", "D1", "E1"])
        assert len(result) == 2, f"Expected 2 ranges, got {result}"
        assert "A1:B1" in result, f"Expected 'A1:B1' in {result}"
        assert "D1:E1" in result, f"Expected 'D1:E1' in {result}"

    def test_single_column_gap(self):
        """Gap in column should create separate ranges."""
        result = combine_cells(["A1", "A2", "A5", "A6"])
        assert len(result) == 2, f"Expected 2 ranges, got {result}"
        assert "A1:A2" in result, f"Expected 'A1:A2' in {result}"
        assert "A5:A6" in result, f"Expected 'A5:A6' in {result}"

    def test_scattered_single_cells(self):
        """Scattered cells should remain individual."""
        result = combine_cells(["A1", "C3", "E5", "G7"])
        assert result == ["A1", "C3", "E5", "G7"], f"Expected individual cells, got {result}"

    def test_l_shape(self):
        """L-shaped cells should form two ranges."""
        cells = ["A1", "A2", "A3", "B3", "C3"]
        result = combine_cells(cells)
        assert len(result) == 2, f"Expected 2 ranges for L-shape, got {result}"

    def test_t_shape(self):
        """T-shaped cells should form multiple ranges."""
        cells = ["B1", "B2", "B3", "A3", "C3"]
        result = combine_cells(cells)
        assert len(result) >= 2, f"Expected at least 2 ranges for T-shape, got {result}"

    def test_double_digit_rows(self):
        """Should handle rows >= 10 correctly."""
        cells = ["A10", "A11", "A12", "B10", "B11", "B12"]
        result = combine_cells(cells)
        assert result == ["A10:B12"], f"Expected ['A10:B12'], got {result}"

    def test_double_letter_columns(self):
        """Should handle columns like AA, AB correctly."""
        cells = ["AA1", "AA2", "AB1", "AB2"]
        result = combine_cells(cells)
        assert result == ["AA1:AB2"], f"Expected ['AA1:AB2'], got {result}"


class TestPrecision:
    """Test that algorithm maintains precision (doesn't create unnecessary ranges)."""

    def test_sparse_corners(self):
        """Four corners of a large area should NOT become one big range."""
        cells = ["A1", "A2", "Z99", "Z100"]
        result = combine_cells(cells)
        # Should be 2 ranges, not 1 giant box
        assert len(result) == 2, f"Expected 2 ranges, got {result}"
        assert "A1:A2" in result, f"Expected 'A1:A2' in {result}"
        assert "Z99:Z100" in result, f"Expected 'Z99:Z100' in {result}"

    def test_diagonal_cells(self):
        """Diagonal cells should remain separate."""
        cells = ["A1", "B2", "C3", "D4"]
        result = combine_cells(cells)
        assert len(result) == 4, f"Expected 4 individual cells, got {result}"

    def test_checkerboard_pattern(self):
        """Checkerboard pattern should not form ranges."""
        cells = ["A1", "A3", "A5", "C1", "C3", "C5", "E1", "E3", "E5"]
        result = combine_cells(cells)
        assert len(result) == 9, f"Expected 9 individual cells, got {result}"

    def test_multiple_small_blocks(self):
        """Multiple small blocks should remain separate."""
        block1 = ["A1", "A2", "B1", "B2"]  # 2x2 at A1
        block2 = ["D1", "D2", "E1", "E2"]  # 2x2 at D1
        block3 = ["A5", "A6", "B5", "B6"]  # 2x2 at A5
        cells = block1 + block2 + block3
        result = combine_cells(cells)
        assert len(result) == 3, f"Expected 3 ranges, got {result}"
        assert "A1:B2" in result, f"Expected 'A1:B2' in {result}"
        assert "D1:E2" in result, f"Expected 'D1:E2' in {result}"
        assert "A5:B6" in result, f"Expected 'A5:B6' in {result}"


class TestPerformance:
    """Test performance on large datasets."""

    def test_large_dense_block(self):
        """Large dense block should be fast and form one range."""
        # 20x20 = 400 cells
        cells = [f"{chr(65 + c)}{r}" for r in range(1, 21) for c in range(20)]

        start = time.time()
        result = combine_cells(cells)
        elapsed = time.time() - start

        print(f"  Large dense block (400 cells): {elapsed*1000:.2f}ms")
        assert len(result) == 1, f"Expected 1 range, got {len(result)} ranges"
        assert elapsed < 0.1, f"Too slow: {elapsed*1000:.2f}ms > 100ms"

    def test_many_individual_cells(self):
        """Many scattered cells should be fast."""
        # 300 scattered cells
        cells = [f"{chr(65 + (i*7) % 26)}{i*3 + 1}" for i in range(300)]

        start = time.time()
        result = combine_cells(cells)
        elapsed = time.time() - start

        print(f"  Scattered cells (300): {elapsed*1000:.2f}ms, {len(result)} ranges")
        assert elapsed < 0.1, f"Too slow: {elapsed*1000:.2f}ms > 100ms"

    def test_many_small_ranges(self):
        """Many small ranges should be fast."""
        # 100 ranges of 3 cells each (vertical), with gaps to keep them separate
        cells = []
        for i in range(100):
            col = chr(65 + (i % 10) * 2)  # A, C, E, G, I, K, M, O, Q, S
            row_start = (i // 10) * 5 + 1  # 1, 6, 11, 16, 21, 26, 31, 36, 41, 46
            cells.extend([f"{col}{row_start}", f"{col}{row_start+1}", f"{col}{row_start+2}"])

        start = time.time()
        result = combine_cells(cells)
        elapsed = time.time() - start

        print(f"  Many small ranges (300 cells -> ~100 ranges): {elapsed*1000:.2f}ms, got {len(result)} ranges")
        # Algorithm might optimize some of these together, so allow some flexibility
        assert len(result) >= 90, f"Expected >=90 ranges, got {len(result)}"
        assert elapsed < 0.1, f"Too slow: {elapsed*1000:.2f}ms > 100ms"

    def test_very_large_sparse_set(self):
        """Very large sparse set (500 cells) should be fast."""
        # 500 scattered cells
        cells = [f"{chr(65 + (i*11) % 26)}{i*5 + 1}" for i in range(500)]

        start = time.time()
        result = combine_cells(cells)
        elapsed = time.time() - start

        print(f"  Very large sparse (500 cells): {elapsed*1000:.2f}ms, {len(result)} ranges")
        assert elapsed < 0.2, f"Too slow: {elapsed*1000:.2f}ms > 200ms"


class TestRealWorldScenarios:
    """Test realistic spreadsheet scenarios."""

    def test_table_with_header(self):
        """Table with header row should form one range."""
        # Header: A1:D1, Data: A2:D5
        cells = []
        for row in range(1, 6):
            for col in range(4):
                cells.append(f"{chr(65 + col)}{row}")

        result = combine_cells(cells)
        assert result == ["A1:D5"], f"Expected ['A1:D5'], got {result}"

    def test_multiple_tables(self):
        """Multiple separate tables should form separate ranges."""
        # Table 1: A1:C5
        table1 = [f"{chr(65 + c)}{r}" for r in range(1, 6) for c in range(3)]
        # Table 2: E1:G5
        table2 = [f"{chr(69 + c)}{r}" for r in range(1, 6) for c in range(3)]
        # Table 3: A8:C12
        table3 = [f"{chr(65 + c)}{r}" for r in range(8, 13) for c in range(3)]

        cells = table1 + table2 + table3
        result = combine_cells(cells)

        assert len(result) == 3, f"Expected 3 ranges, got {len(result)}: {result}"
        assert "A1:C5" in result, f"Expected 'A1:C5' in {result}"
        assert "E1:G5" in result, f"Expected 'E1:G5' in {result}"
        assert "A8:C12" in result, f"Expected 'A8:C12' in {result}"

    def test_pivot_table_layout(self):
        """Pivot table with row/column headers."""
        # Row headers: A2:A5
        # Column headers: B1:D1
        # Data: B2:D5
        cells = (
            [f"A{r}" for r in range(2, 6)] +  # Row headers
            [f"{chr(66 + c)}1" for c in range(3)] +  # Column headers
            [f"{chr(66 + c)}{r}" for r in range(2, 6) for c in range(3)]  # Data
        )

        result = combine_cells(cells)
        # Should recognize this as one connected region
        assert len(result) <= 2, f"Expected <=2 ranges for pivot table, got {len(result)}: {result}"

    def test_timeline_data(self):
        """Timeline/Gantt chart style data with gaps."""
        # Project 1: B1:E1
        # Project 2: D2:G2
        # Project 3: A3:C3
        cells = (
            ["B1", "C1", "D1", "E1"] +
            ["D2", "E2", "F2", "G2"] +
            ["A3", "B3", "C3"]
        )

        result = combine_cells(cells)
        assert len(result) == 3, f"Expected 3 ranges, got {len(result)}: {result}"


def run_all_tests():
    """Run all test classes."""
    test_classes = [
        TestColToIndex,
        TestIndexToCol,
        TestBasicFunctionality,
        TestEdgeCases,
        TestPrecision,
        TestPerformance,
        TestRealWorldScenarios,
    ]

    total_tests = 0
    passed_tests = 0
    failed_tests = []

    for test_class in test_classes:
        print(f"\n{'='*70}")
        print(f"Running {test_class.__name__}")
        print(f"{'='*70}")

        test_instance = test_class()
        test_methods = [m for m in dir(test_instance) if m.startswith("test_")]

        for method_name in test_methods:
            total_tests += 1
            try:
                method = getattr(test_instance, method_name)
                method()
                print(f"✓ {method_name}")
                passed_tests += 1
            except AssertionError as e:
                print(f"✗ {method_name}: {e}")
                failed_tests.append((test_class.__name__, method_name, str(e)))
            except Exception as e:
                print(f"✗ {method_name}: EXCEPTION: {e}")
                failed_tests.append((test_class.__name__, method_name, f"Exception: {e}"))

    # Summary
    print(f"\n{'='*70}")
    print("TEST SUMMARY")
    print(f"{'='*70}")
    print(f"Total tests: {total_tests}")
    print(f"Passed: {passed_tests}")
    print(f"Failed: {len(failed_tests)}")

    if failed_tests:
        print("\nFailed tests:")
        for class_name, method_name, error in failed_tests:
            print(f"  - {class_name}.{method_name}")
            print(f"    {error}")
        return False
    else:
        print("\n🎉 All tests passed!")
        return True


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
