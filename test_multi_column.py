"""
test_multi_column.py
Testy dla funkcjonalności wyszukiwania wielu kolumn o tej samej nazwie.

Testy jednostkowe weryfikujące:
- Znajdowanie wszystkich kolumn o tej samej nazwie w jednym arkuszu
- Dopasowanie case-insensitive i z ignorowaniem białych znaków
- Obsługę wielu kolumn o tej samej nazwie w różnych zakładkach
"""

import unittest
from sheets_search import (
    find_all_column_indices_by_name,
    normalize_header_name,
    find_column_index_by_name,
)


class TestMultiColumnSearch(unittest.TestCase):
    """Testy funkcjonalności wyszukiwania wielu kolumn."""

    def test_find_all_column_indices_empty_header(self):
        """Test: pusta lista nagłówków zwraca pustą listę indeksów."""
        result = find_all_column_indices_by_name([], "Test")
        self.assertEqual(result, [])

    def test_find_all_column_indices_no_match(self):
        """Test: brak dopasowania zwraca pustą listę."""
        headers = ["Kolumna A", "Kolumna B", "Kolumna C"]
        result = find_all_column_indices_by_name(headers, "Nie ma")
        self.assertEqual(result, [])

    def test_find_all_column_indices_single_match(self):
        """Test: pojedyncze dopasowanie zwraca listę z jednym indeksem."""
        headers = ["Kolumna A", "Test", "Kolumna C"]
        result = find_all_column_indices_by_name(headers, "Test")
        self.assertEqual(result, [1])

    def test_find_all_column_indices_multiple_matches(self):
        """Test: wiele dopasowań zwraca listę wszystkich indeksów."""
        headers = ["Test", "Kolumna B", "Test", "Kolumna D", "test"]
        result = find_all_column_indices_by_name(headers, "Test")
        # Powinny znaleźć indeksy 0, 2, 4 (case-insensitive)
        self.assertEqual(result, [0, 2, 4])

    def test_find_all_column_indices_case_insensitive(self):
        """Test: wyszukiwanie jest niewrażliwe na wielkość liter."""
        headers = ["TEST", "test", "TeSt", "other"]
        result = find_all_column_indices_by_name(headers, "test")
        self.assertEqual(result, [0, 1, 2])

    def test_find_all_column_indices_whitespace_handling(self):
        """Test: ignorowanie białych znaków w dopasowaniu."""
        headers = ["Test Column", " Test Column ", "Test  Column", "Other"]
        result = find_all_column_indices_by_name(headers, "Test Column")
        # Wszystkie trzy pierwsze powinny pasować (normalizacja białych znaków)
        self.assertEqual(result, [0, 1, 2])

    def test_find_all_column_indices_underscore_normalization(self):
        """Test: podkreślniki są zamieniane na spacje."""
        headers = ["test_column", "test column", "testcolumn"]
        result = find_all_column_indices_by_name(headers, "test_column")
        # "test_column" i "test column" powinny pasować
        self.assertEqual(result, [0, 1])

    def test_find_all_column_indices_with_none_values(self):
        """Test: obsługa None w nagłówkach."""
        headers = ["Test", None, "Test", "Other", None]
        result = find_all_column_indices_by_name(headers, "Test")
        self.assertEqual(result, [0, 2])

    def test_normalize_header_name_basic(self):
        """Test: normalizacja nazw nagłówków."""
        self.assertEqual(normalize_header_name("Test"), "test")
        self.assertEqual(normalize_header_name("TEST"), "test")
        self.assertEqual(normalize_header_name(" Test "), "test")
        self.assertEqual(normalize_header_name("Test_Column"), "test column")
        self.assertEqual(normalize_header_name("Test  Multiple  Spaces"), "test multiple spaces")

    def test_normalize_header_name_none(self):
        """Test: normalizacja None zwraca pusty string."""
        self.assertEqual(normalize_header_name(None), "")

    def test_normalize_header_name_numeric(self):
        """Test: normalizacja wartości numerycznych."""
        self.assertEqual(normalize_header_name(123), "123")
        self.assertEqual(normalize_header_name(45.67), "45.67")

    def test_backward_compatibility_single_column(self):
        """Test: kompatybilność wsteczna - find_column_index_by_name zwraca pierwszy indeks."""
        headers = ["Test", "Other", "Test"]
        # Stara funkcja powinna zwrócić pierwszy indeks (0)
        result_old = find_column_index_by_name(headers, "Test")
        self.assertEqual(result_old, 0)
        
        # Nowa funkcja powinna zwrócić wszystkie indeksy [0, 2]
        result_new = find_all_column_indices_by_name(headers, "Test")
        self.assertEqual(result_new, [0, 2])

    def test_multiple_columns_different_positions(self):
        """Test: kolumny o tej samej nazwie w różnych pozycjach."""
        headers = ["A", "Numer zlecenia", "B", "C", "numer_zlecenia", "D", "NUMER ZLECENIA"]
        result = find_all_column_indices_by_name(headers, "Numer zlecenia")
        # Powinny znaleźć indeksy 1, 4, 6 (wszystkie warianty "numer zlecenia")
        self.assertEqual(result, [1, 4, 6])

    def test_empty_column_name(self):
        """Test: pusta nazwa kolumny zwraca pustą listę."""
        headers = ["A", "B", "C"]
        result = find_all_column_indices_by_name(headers, "")
        self.assertEqual(result, [])

    def test_none_column_name(self):
        """Test: None jako nazwa kolumny zwraca pustą listę."""
        headers = ["A", "B", "C"]
        result = find_all_column_indices_by_name(headers, None)
        self.assertEqual(result, [])


if __name__ == "__main__":
    # Uruchom testy
    unittest.main(verbosity=2)
