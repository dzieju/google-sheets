"""
test_header_rows.py
Testy dla funkcjonalności wielu wierszy nagłówkowych (header rows).

Testy jednostkowe weryfikujące:
- Parsowanie konfiguracji header rows (np. "1", "1,2", "1, 2, 3")
- Łączenie wartości z wielu wierszy nagłówkowych dla każdej kolumny
- Wykrywanie nagłówków z wykorzystaniem multiple header rows
- Kompatybilność wsteczną (domyślnie wiersz 1)
"""

import unittest
from sheets_search import (
    parse_header_rows,
    combine_header_values,
    detect_header_row,
)


class TestHeaderRows(unittest.TestCase):
    """Testy funkcjonalności wielu wierszy nagłówkowych."""

    # -------------------- Testy parse_header_rows --------------------
    
    def test_parse_header_rows_none(self):
        """Test: None jako input zwraca [0] (default: wiersz 1)."""
        result = parse_header_rows(None)
        self.assertEqual(result, [0])

    def test_parse_header_rows_empty(self):
        """Test: pusty string zwraca [0] (default: wiersz 1)."""
        result = parse_header_rows("")
        self.assertEqual(result, [0])
        
        result = parse_header_rows("   ")
        self.assertEqual(result, [0])

    def test_parse_header_rows_single_value(self):
        """Test: pojedyncza wartość "1" zwraca [0]."""
        result = parse_header_rows("1")
        self.assertEqual(result, [0])
        
    def test_parse_header_rows_single_value_2(self):
        """Test: pojedyncza wartość "2" zwraca [1]."""
        result = parse_header_rows("2")
        self.assertEqual(result, [1])

    def test_parse_header_rows_multiple_values(self):
        """Test: wiele wartości "1,2" zwraca [0, 1]."""
        result = parse_header_rows("1,2")
        self.assertEqual(result, [0, 1])

    def test_parse_header_rows_multiple_values_with_spaces(self):
        """Test: wartości z białymi znakami "1, 2, 3" zwraca [0, 1, 2]."""
        result = parse_header_rows("1, 2, 3")
        self.assertEqual(result, [0, 1, 2])

    def test_parse_header_rows_invalid_values(self):
        """Test: nieprawidłowe wartości są pomijane."""
        result = parse_header_rows("1, abc, 2")
        self.assertEqual(result, [0, 1])
        
        result = parse_header_rows("abc, def")
        self.assertEqual(result, [0])  # Brak prawidłowych wartości -> default [0]

    def test_parse_header_rows_zero_or_negative(self):
        """Test: wartości <= 0 są pomijane."""
        result = parse_header_rows("0, 1, -1, 2")
        self.assertEqual(result, [0, 1])

    # -------------------- Testy combine_header_values --------------------

    def test_combine_header_values_empty(self):
        """Test: pusta lista wierszy zwraca pustą listę."""
        result = combine_header_values([], [0])
        self.assertEqual(result, [])

    def test_combine_header_values_single_row(self):
        """Test: pojedynczy wiersz nagłówkowy."""
        values = [["Name", "Age", "City"]]
        result = combine_header_values(values, [0])
        # Normalizacja: lowercase
        self.assertEqual(result, ["name", "age", "city"])

    def test_combine_header_values_two_rows(self):
        """Test: dwa wiersze nagłówkowe są łączone spacją."""
        values = [
            ["First", "Age", "Home"],
            ["Last", "Years", "City"]
        ]
        result = combine_header_values(values, [0, 1])
        # Łączenie: "First Last", "Age Years", "Home City" -> lowercase
        self.assertEqual(result, ["first last", "age years", "home city"])

    def test_combine_header_values_three_rows(self):
        """Test: trzy wiersze nagłówkowe."""
        values = [
            ["A", "B", "C"],
            ["1", "2", "3"],
            ["X", "Y", "Z"]
        ]
        result = combine_header_values(values, [0, 1, 2])
        self.assertEqual(result, ["a 1 x", "b 2 y", "c 3 z"])

    def test_combine_header_values_with_none(self):
        """Test: wartości None są pomijane."""
        values = [
            ["Name", None, "City"],
            [None, "Age", "Location"]
        ]
        result = combine_header_values(values, [0, 1])
        # "Name", "Age", "City Location" -> lowercase
        self.assertEqual(result, ["name", "age", "city location"])

    def test_combine_header_values_with_empty_strings(self):
        """Test: puste stringi są pomijane."""
        values = [
            ["Name", "", "City"],
            ["", "Age", ""]
        ]
        result = combine_header_values(values, [0, 1])
        self.assertEqual(result, ["name", "age", "city"])

    def test_combine_header_values_different_column_counts(self):
        """Test: wiersze o różnej liczbie kolumn."""
        values = [
            ["A", "B", "C"],
            ["1", "2"]
        ]
        result = combine_header_values(values, [0, 1])
        # Kolumna C jest tylko w wierszu 0
        self.assertEqual(result, ["a 1", "b 2", "c"])

    def test_combine_header_values_normalization(self):
        """Test: normalizacja - podkreślenia i wielokrotne spacje."""
        values = [
            ["First_Name", "Age  In", "Home"],
            ["Last  Name", "Years", "City_Location"]
        ]
        result = combine_header_values(values, [0, 1])
        # Normalizacja zamienia '_' na ' ' i redukuje wielokrotne spacje
        self.assertEqual(result, ["first name last name", "age in years", "home city location"])

    # -------------------- Testy integracji detect_header_row z header_row_indices --------------------

    def test_detect_header_row_with_single_index(self):
        """Test: detect_header_row z podanym pojedynczym indeksem."""
        values = [
            ["Name", "Age", "City"],
            ["Alice", "30", "NYC"],
            ["Bob", "25", "LA"]
        ]
        header_row_idx, header_row, start_row = detect_header_row(values, header_row_indices=[0])
        self.assertEqual(header_row_idx, [0])
        self.assertEqual(header_row, ["name", "age", "city"])
        self.assertEqual(start_row, 1)

    def test_detect_header_row_with_multiple_indices(self):
        """Test: detect_header_row z podanymi wieloma indeksami."""
        values = [
            ["First", "Age", "Home"],
            ["Last", "Years", "City"],
            ["Alice", "30", "NYC"],
            ["Bob", "25", "LA"]
        ]
        header_row_idx, header_row, start_row = detect_header_row(values, header_row_indices=[0, 1])
        self.assertEqual(header_row_idx, [0, 1])
        self.assertEqual(header_row, ["first last", "age years", "home city"])
        self.assertEqual(start_row, 2)  # Dane zaczynają się od wiersza 2 (indeks 2)

    def test_detect_header_row_with_no_indices_auto_detection(self):
        """Test: detect_header_row bez indeksów - auto-detekcja."""
        values = [
            ["Name", "Age", "City"],
            ["Alice", "30", "NYC"],
            ["Bob", "25", "LA"]
        ]
        header_row_idx, header_row, start_row = detect_header_row(values, header_row_indices=None)
        # Auto-detekcja powinna wybrać wiersz 0
        self.assertEqual(header_row_idx, 0)
        self.assertEqual(header_row, ["Name", "Age", "City"])
        self.assertEqual(start_row, 1)

    def test_detect_header_row_with_indices_row_2_and_3(self):
        """Test: detect_header_row z indeksami [1, 2] (wiersze 2 i 3)."""
        values = [
            ["Skip this row"],
            ["First", "Age", "Home"],
            ["Last", "Years", "City"],
            ["Alice", "30", "NYC"]
        ]
        header_row_idx, header_row, start_row = detect_header_row(values, header_row_indices=[1, 2])
        self.assertEqual(header_row_idx, [1, 2])
        self.assertEqual(header_row, ["first last", "age years", "home city"])
        self.assertEqual(start_row, 3)

    def test_backward_compatibility_no_header_row_indices(self):
        """Test: kompatybilność wsteczna - brak header_row_indices działa jak dotychczas."""
        values = [
            ["Name", "Age", "City"],
            ["Alice", "30", "NYC"]
        ]
        # Bez header_row_indices
        result_old = detect_header_row(values, search_column_name=None, header_row_indices=None)
        # Z header_row_indices=[0] (domyślne)
        result_new = detect_header_row(values, search_column_name=None, header_row_indices=[0])
        
        # Oba powinny zwrócić nagłówki z wiersza 0, ale w różnych formatach
        # result_old zwraca pojedynczy indeks 0, result_new zwraca listę [0]
        self.assertEqual(result_old[2], result_new[2])  # start_row powinien być taki sam
        # Header row w result_new jest znormalizowany (lowercase)
        self.assertEqual(result_new[1], ["name", "age", "city"])


if __name__ == "__main__":
    # Uruchom testy
    unittest.main(verbosity=2)
