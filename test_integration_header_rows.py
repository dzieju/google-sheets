"""
test_integration_header_rows.py
Testy integracyjne dla funkcjonalności wielu wierszy nagłówkowych.

Testy weryfikujące działanie wyszukiwania z wykorzystaniem wielu wierszy nagłówkowych
w kontekście całego systemu (bez rzeczywistego API Google Sheets).
"""

import unittest
from sheets_search import (
    parse_header_rows,
    combine_header_values,
    detect_header_row,
    find_all_column_indices_by_name,
    matches_ignore_pattern,
    parse_ignore_patterns,
)


class TestHeaderRowsIntegration(unittest.TestCase):
    """Testy integracyjne funkcjonalności wielu wierszy nagłówkowych."""

    def test_search_with_multi_header_rows_and_ignore(self):
        """
        Test: wyszukiwanie z wieloma wierszami nagłówkowymi i ignorowaniem kolumn.
        
        Symuluje scenariusz gdzie:
        - Wiersz 1 zawiera: ["First", "Age", "Home", "Temp"]
        - Wiersz 2 zawiera: ["Last", "Years", "City", "Data"]
        - Kolumna "Temp Data" powinna być ignorowana (wzorzec "temp*")
        - Wyszukujemy kolumnę "First Last"
        """
        values = [
            ["First", "Age", "Home", "Temp"],
            ["Last", "Years", "City", "Data"],
            ["Alice", "30", "NYC", "123"],
            ["Bob", "25", "LA", "456"]
        ]
        
        # Parse header rows configuration
        header_row_indices = parse_header_rows("1,2")
        self.assertEqual(header_row_indices, [0, 1])
        
        # Detect header row and combine headers
        header_row_idx, header_row, start_row = detect_header_row(
            values, 
            search_column_name=None,
            header_row_indices=header_row_indices
        )
        
        self.assertEqual(header_row_idx, [0, 1])
        self.assertEqual(header_row, ["first last", "age years", "home city", "temp data"])
        self.assertEqual(start_row, 2)  # Dane zaczynają się od wiersza 2 (indeks 2)
        
        # Parse ignore patterns
        ignore_patterns = parse_ignore_patterns("temp*")
        self.assertEqual(ignore_patterns, ["temp*"])
        
        # Check which columns should be found
        # Szukamy kolumny "First Last" - powinna być znaleziona
        target_indices = find_all_column_indices_by_name(
            header_row, 
            "First Last", 
            ignore_patterns
        )
        self.assertEqual(target_indices, [0])  # Kolumna 0 pasuje
        
        # Sprawdź czy kolumna "Temp Data" jest ignorowana
        temp_indices = find_all_column_indices_by_name(
            header_row,
            "Temp Data",
            ignore_patterns
        )
        self.assertEqual(temp_indices, [])  # Kolumna 3 powinna być ignorowana

    def test_search_with_single_header_row(self):
        """
        Test: wyszukiwanie z pojedynczym wierszem nagłówkowym (kompatybilność wsteczna).
        """
        values = [
            ["Name", "Age", "City"],
            ["Alice", "30", "NYC"],
            ["Bob", "25", "LA"]
        ]
        
        # Parse header rows configuration - default "1"
        header_row_indices = parse_header_rows("1")
        self.assertEqual(header_row_indices, [0])
        
        # Detect header row
        header_row_idx, header_row, start_row = detect_header_row(
            values,
            search_column_name=None,
            header_row_indices=header_row_indices
        )
        
        self.assertEqual(header_row_idx, [0])
        self.assertEqual(header_row, ["name", "age", "city"])
        self.assertEqual(start_row, 1)

    def test_multi_header_with_query_matching(self):
        """
        Test: dopasowanie zapytania do połączonego nagłówka.
        
        Scenario: użytkownik szuka kolumny o nazwie "product name" 
        a w arkuszu mamy:
        - Wiersz 1: ["Product", "Price", "Stock"]
        - Wiersz 2: ["Name", "USD", "Quantity"]
        """
        values = [
            ["Product", "Price", "Stock"],
            ["Name", "USD", "Quantity"],
            ["Laptop", "999", "10"],
            ["Mouse", "25", "50"]
        ]
        
        # Parse header rows
        header_row_indices = parse_header_rows("1,2")
        
        # Detect and combine headers
        header_row_idx, header_row, start_row = detect_header_row(
            values,
            header_row_indices=header_row_indices
        )
        
        # Combined headers: "product name", "price usd", "stock quantity"
        self.assertEqual(header_row, ["product name", "price usd", "stock quantity"])
        
        # Find column matching "Product Name"
        indices = find_all_column_indices_by_name(header_row, "Product Name")
        self.assertEqual(indices, [0])

    def test_multi_header_with_wildcards_and_ignore(self):
        """
        Test: kombinacja wielu wierszy nagłówkowych i ignorowania z wildcardami.
        """
        values = [
            ["Test", "Prod", "Debug", "Final"],
            ["Data", "Value", "Info", "Result"],
            ["1", "2", "3", "4"]
        ]
        
        header_row_indices = parse_header_rows("1,2")
        header_row_idx, header_row, start_row = detect_header_row(
            values,
            header_row_indices=header_row_indices
        )
        
        # Combined: "test data", "prod value", "debug info", "final result"
        self.assertEqual(header_row, ["test data", "prod value", "debug info", "final result"])
        
        # Ignore patterns: "*debug*", "test*"
        ignore_patterns = parse_ignore_patterns("*debug*, test*")
        
        # Check which columns are ignored
        self.assertTrue(matches_ignore_pattern("test data", ignore_patterns))  # test* matches
        self.assertTrue(matches_ignore_pattern("debug info", ignore_patterns))  # *debug* matches
        self.assertFalse(matches_ignore_pattern("prod value", ignore_patterns))
        self.assertFalse(matches_ignore_pattern("final result", ignore_patterns))
        
        # Find non-ignored columns matching "prod value"
        indices = find_all_column_indices_by_name(header_row, "Prod Value", ignore_patterns)
        self.assertEqual(indices, [1])

    def test_empty_header_rows_defaults_to_one(self):
        """Test: puste pole Header rows domyślnie używa wiersza 1."""
        header_row_indices = parse_header_rows("")
        self.assertEqual(header_row_indices, [0])
        
        header_row_indices = parse_header_rows(None)
        self.assertEqual(header_row_indices, [0])

    def test_invalid_header_rows_defaults_to_one(self):
        """Test: nieprawidłowe wartości w Header rows domyślnie używają wiersza 1."""
        header_row_indices = parse_header_rows("abc")
        self.assertEqual(header_row_indices, [0])
        
        header_row_indices = parse_header_rows("0, -1")
        self.assertEqual(header_row_indices, [0])


if __name__ == "__main__":
    unittest.main(verbosity=2)
