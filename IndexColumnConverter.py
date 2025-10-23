import string
import re


class IndexColumnConverter:

    def __init__(self):
        return

    # Converts index to column letter; courtesy of https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
    def parse_colindex(self, num):

        # Modified divmod function for Excel; courtesy of https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
        def divmod_excel(n):
            a, b = divmod(n, 26)
            if b == 0:
                return a - 1, b + 26
            return a, b

        chars = []
        while num > 0:
            num, d = divmod_excel(num)
            chars.append(string.ascii_uppercase[d - 1])
        return "".join(reversed(chars))

    # Converts column letter to index (1-based)
    def parse_cellindex(self, col_str):
        """
        Convert column letter(s) to 1-based index.

        Args:
            col_str: Column letter(s) like 'A', 'B', 'AA', etc.

        Returns:
            1-based column index (A=1, B=2, ..., Z=26, AA=27, etc.)

        Examples:
            'A' -> 1
            'B' -> 2
            'Z' -> 26
            'AA' -> 27
        """
        result = 0
        for char in col_str.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    # Parse cell coordinate into column and row parts
    def parse_cell(self, cell_str):
        """
        Parse cell coordinate string into column and row parts.

        Args:
            cell_str: Cell coordinate like 'A1', 'B5', 'AA100', etc.

        Returns:
            Tuple of (column_str, row_str) or None if invalid format

        Examples:
            'A1' -> ('A', '1')
            'B5' -> ('B', '5')
            'AA100' -> ('AA', '100')
        """
        match = re.match(r'^([A-Z]+)(\d+)$', cell_str.upper())
        if match:
            return match.group(1), match.group(2)
        return None
