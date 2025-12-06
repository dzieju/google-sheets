"""
test_ignore_patterns.py
Testy dla funkcjonalności pola "Ignoruj" (ignore patterns).

Testy jednostkowe weryfikujące:
- Parsowanie wielu wartości z pola Ignoruj
- Dopasowanie wzorców z wildcardami (*, *pattern, pattern*, *pattern*)
- Ignorowanie kolumn pasujących do wzorców
- Priorytet Ignoruj nad Zapytaniem (ignore overrides query)
- Kompatybilność wsteczną (puste pole Ignoruj)
"""

import unittest
from sheets_search import (
    parse_ignore_patterns,
    matches_ignore_pattern,
    find_all_column_indices_by_name,
)


class TestIgnorePatterns(unittest.TestCase):
    """Testy funkcjonalności ignorowania kolumn."""

    # -------------------- Testy parse_ignore_patterns --------------------
    
    def test_parse_ignore_patterns_none(self):
        """Test: None jako input zwraca pustą listę."""
        result = parse_ignore_patterns(None)
        self.assertEqual(result, [])

    def test_parse_ignore_patterns_empty(self):
        """Test: pusty string zwraca pustą listę."""
        result = parse_ignore_patterns("")
        self.assertEqual(result, [])
        
        result = parse_ignore_patterns("   ")
        self.assertEqual(result, [])

    def test_parse_ignore_patterns_single_value(self):
        """Test: pojedyncza wartość."""
        result = parse_ignore_patterns("temp")
        self.assertEqual(result, ["temp"])

    def test_parse_ignore_patterns_comma_separated(self):
        """Test: wartości oddzielone przecinkami."""
        result = parse_ignore_patterns("temp, test, debug")
        self.assertEqual(result, ["temp", "test", "debug"])

    def test_parse_ignore_patterns_semicolon_separated(self):
        """Test: wartości oddzielone średnikami."""
        result = parse_ignore_patterns("temp; test; debug")
        self.assertEqual(result, ["temp", "test", "debug"])

    def test_parse_ignore_patterns_newline_separated(self):
        """Test: wartości oddzielone nowymi liniami."""
        result = parse_ignore_patterns("temp\ntest\ndebug")
        self.assertEqual(result, ["temp", "test", "debug"])

    def test_parse_ignore_patterns_mixed_separators(self):
        """Test: mieszane separatory (przecinki, średniki, nowe linie)."""
        result = parse_ignore_patterns("temp, test; debug\nold")
        self.assertEqual(result, ["temp", "test", "debug", "old"])

    def test_parse_ignore_patterns_with_whitespace(self):
        """Test: wartości z białymi znakami są trimowane."""
        result = parse_ignore_patterns("  temp  ,   test   ,  debug  ")
        self.assertEqual(result, ["temp", "test", "debug"])

    def test_parse_ignore_patterns_lowercase(self):
        """Test: wartości są konwertowane na lowercase."""
        result = parse_ignore_patterns("TEMP, Test, DeBuG")
        self.assertEqual(result, ["temp", "test", "debug"])

    def test_parse_ignore_patterns_with_wildcards(self):
        """Test: wzorce z wildcardami są zachowane."""
        result = parse_ignore_patterns("temp*, *test, *debug*")
        self.assertEqual(result, ["temp*", "*test", "*debug*"])

    def test_parse_ignore_patterns_empty_values_filtered(self):
        """Test: puste wartości są pomijane."""
        result = parse_ignore_patterns("temp, , test, , debug")
        self.assertEqual(result, ["temp", "test", "debug"])

    # -------------------- Testy matches_ignore_pattern --------------------

    def test_matches_ignore_pattern_empty_patterns(self):
        """Test: pusta lista wzorców zwraca False."""
        result = matches_ignore_pattern("test", [])
        self.assertFalse(result)

    def test_matches_ignore_pattern_exact_match(self):
        """Test: dokładne dopasowanie."""
        result = matches_ignore_pattern("temp", ["temp"])
        self.assertTrue(result)

    def test_matches_ignore_pattern_exact_match_case_insensitive(self):
        """Test: dokładne dopasowanie case-insensitive."""
        result = matches_ignore_pattern("TEMP", ["temp"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("temp", ["TEMP"])
        self.assertTrue(result)

    def test_matches_ignore_pattern_exact_match_with_whitespace(self):
        """Test: dopasowanie z ignorowaniem białych znaków."""
        result = matches_ignore_pattern("  temp  ", ["temp"])
        self.assertTrue(result)

    def test_matches_ignore_pattern_no_match(self):
        """Test: brak dopasowania."""
        result = matches_ignore_pattern("production", ["temp", "test", "debug"])
        self.assertFalse(result)

    def test_matches_ignore_pattern_prefix_wildcard(self):
        """Test: wildcard na końcu (pattern*) - startsWith."""
        result = matches_ignore_pattern("temporary", ["temp*"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("temp_col", ["temp*"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("notemp", ["temp*"])
        self.assertFalse(result)

    def test_matches_ignore_pattern_suffix_wildcard(self):
        """Test: wildcard na początku (*pattern) - endsWith."""
        result = matches_ignore_pattern("old_temp", ["*temp"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("debug_temp", ["*temp"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("temporary", ["*temp"])
        self.assertFalse(result)

    def test_matches_ignore_pattern_contains_wildcard(self):
        """Test: wildcard z obu stron (*pattern*) - contains."""
        result = matches_ignore_pattern("old_debug_temp", ["*debug*"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("debug", ["*debug*"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("production", ["*debug*"])
        self.assertFalse(result)

    def test_matches_ignore_pattern_multiple_patterns(self):
        """Test: dopasowanie do jednego z wielu wzorców."""
        patterns = ["temp*", "*debug*", "old"]
        
        self.assertTrue(matches_ignore_pattern("temporary", patterns))
        self.assertTrue(matches_ignore_pattern("debug_mode", patterns))
        self.assertTrue(matches_ignore_pattern("old", patterns))
        self.assertFalse(matches_ignore_pattern("production", patterns))

    def test_matches_ignore_pattern_normalized_header(self):
        """Test: dopasowanie z normalizacją nagłówka (podkreślniki, spacje)."""
        # normalize_header_name zamienia '_' na ' ' i redukuje spacje
        result = matches_ignore_pattern("temp_column", ["temp column"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("temp  column", ["temp column"])
        self.assertTrue(result)

    def test_matches_ignore_pattern_empty_header(self):
        """Test: pusty nagłówek zwraca False."""
        result = matches_ignore_pattern("", ["temp"])
        self.assertFalse(result)

    # -------------------- Testy integracji z find_all_column_indices_by_name --------------------

    def test_find_all_column_indices_no_ignore(self):
        """Test: bez wzorców ignorowania - zwraca wszystkie pasujące kolumny."""
        headers = ["Test", "Other", "Test", "Test"]
        result = find_all_column_indices_by_name(headers, "Test", ignore_patterns=None)
        self.assertEqual(result, [0, 2, 3])

    def test_find_all_column_indices_with_exact_ignore(self):
        """Test: ignorowanie dokładnego dopasowania."""
        headers = ["Test", "Other", "Test", "Test"]
        ignore = parse_ignore_patterns("test")
        result = find_all_column_indices_by_name(headers, "Test", ignore_patterns=ignore)
        # Wszystkie kolumny "Test" są ignorowane
        self.assertEqual(result, [])

    def test_find_all_column_indices_with_wildcard_ignore(self):
        """Test: ignorowanie z wildcardem."""
        headers = ["Test", "Test_Old", "Test_New", "Other"]
        ignore = parse_ignore_patterns("*old")
        result = find_all_column_indices_by_name(headers, "Test", ignore_patterns=ignore)
        # Tylko indeks 1 (Test_Old) powinien być ignorowany
        self.assertEqual(result, [0])
        
        # Inny przykład - ignorowanie wszystkich kolumn zaczynających się na "test"
        headers2 = ["Test", "Test_Column", "Prod_Test", "Other"]
        ignore2 = parse_ignore_patterns("test*")
        result2 = find_all_column_indices_by_name(headers2, "Test", ignore_patterns=ignore2)
        # Wszystkie kolumny "Test" pasują do "test*" więc są ignorowane
        self.assertEqual(result2, [])

    def test_find_all_column_indices_partial_ignore(self):
        """Test: tylko część kolumn jest ignorowana."""
        headers = ["Numer_Zlecenia", "Numer_Zlecenia_Old", "Numer_Zlecenia"]
        ignore = parse_ignore_patterns("*old")
        result = find_all_column_indices_by_name(headers, "Numer_Zlecenia", ignore_patterns=ignore)
        # Indeksy 0 i 2 powinny być zwrócone (1 jest ignorowany)
        self.assertEqual(result, [0, 2])

    def test_find_all_column_indices_multiple_ignore_patterns(self):
        """Test: wiele wzorców ignorowania."""
        headers = ["Test", "Test_Old", "Test_Debug", "Test", "Other"]
        ignore = parse_ignore_patterns("*old, *debug")
        result = find_all_column_indices_by_name(headers, "Test", ignore_patterns=ignore)
        # Indeksy 0 i 3 powinny być zwrócone (1 i 2 są ignorowane)
        self.assertEqual(result, [0, 3])

    def test_backward_compatibility_empty_ignore(self):
        """Test: kompatybilność wsteczna - puste ignore_patterns działa jak None."""
        headers = ["Test", "Other", "Test"]
        result_none = find_all_column_indices_by_name(headers, "Test", ignore_patterns=None)
        result_empty = find_all_column_indices_by_name(headers, "Test", ignore_patterns=[])
        self.assertEqual(result_none, result_empty)
        self.assertEqual(result_none, [0, 2])


if __name__ == "__main__":
    # Uruchom testy
    unittest.main(verbosity=2)
