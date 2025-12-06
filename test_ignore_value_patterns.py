"""
test_ignore_value_patterns.py
Testy dla nowej funkcjonalności ignorowania wartości komórek (nie tylko nagłówków).

Testy jednostkowe weryfikujące:
- Parsowanie wzorców pozostaje bez zmian (z test_ignore_patterns.py)
- Nowa semantyka dopasowania: pattern bez '*' = substring match
- Ignorowanie dopasowanych wartości komórek zawierających ignore patterns
- Przypadek użytkownika: "https" w Ignoruj pomija URL-e
- Zachowanie wildcard patterns (*pattern, pattern*, *pattern*)
- Kompatybilność wsteczna
"""

import unittest
from sheets_search import (
    parse_ignore_patterns,
    matches_ignore_pattern,
    matches_ignore_value,
)


class TestIgnoreValuePatterns(unittest.TestCase):
    """Testy funkcjonalności ignorowania wartości komórek."""

    # -------------------- Testy matches_ignore_value (nowa funkcja) --------------------
    
    def test_matches_ignore_value_none_patterns(self):
        """Test: None jako patterns zwraca False."""
        result = matches_ignore_value("some value", None)
        self.assertFalse(result)
    
    def test_matches_ignore_value_empty_patterns(self):
        """Test: pusta lista wzorców zwraca False."""
        result = matches_ignore_value("some value", [])
        self.assertFalse(result)
    
    def test_matches_ignore_value_empty_value(self):
        """Test: pusta wartość zwraca False."""
        result = matches_ignore_value("", ["https"])
        self.assertFalse(result)
        
        result = matches_ignore_value(None, ["https"])
        self.assertFalse(result)
    
    def test_matches_ignore_value_substring_match(self):
        """Test: pattern bez '*' = substring match (case-insensitive)."""
        # Główny przypadek użytkownika: "https" ignoruje URL-e
        result = matches_ignore_value("https://example.com/path?param=value", ["https"])
        self.assertTrue(result)
        
        result = matches_ignore_value("HTTPS://EXAMPLE.COM", ["https"])
        self.assertTrue(result)
        
        result = matches_ignore_value("some https:// text", ["https"])
        self.assertTrue(result)
        
        # Wartość nie zawierająca wzorca
        result = matches_ignore_value("some text without protocol", ["https"])
        self.assertFalse(result)
        
        # HTTP (bez 's') nie pasuje do "https"
        result = matches_ignore_value("HTTP://EXAMPLE.COM", ["https"])
        self.assertFalse(result)
    
    def test_matches_ignore_value_case_insensitive(self):
        """Test: dopasowanie case-insensitive."""
        patterns = ["test"]
        
        self.assertTrue(matches_ignore_value("TEST", patterns))
        self.assertTrue(matches_ignore_value("Test", patterns))
        self.assertTrue(matches_ignore_value("test", patterns))
        self.assertTrue(matches_ignore_value("This is a TEST value", patterns))
    
    def test_matches_ignore_value_with_trim(self):
        """Test: wartości są trimowane przed porównaniem."""
        result = matches_ignore_value("  https://example.com  ", ["https"])
        self.assertTrue(result)
        
        result = matches_ignore_value("value with spaces", ["  with  "])
        self.assertTrue(result)
    
    def test_matches_ignore_value_prefix_wildcard(self):
        """Test: wildcard na końcu (pattern*) - startsWith."""
        patterns = ["http*"]
        
        self.assertTrue(matches_ignore_value("https://example.com", patterns))
        self.assertTrue(matches_ignore_value("http://example.com", patterns))
        self.assertFalse(matches_ignore_value("ftp://example.com", patterns))
    
    def test_matches_ignore_value_suffix_wildcard(self):
        """Test: wildcard na początku (*pattern) - endsWith."""
        patterns = ["*.pdf"]
        
        self.assertTrue(matches_ignore_value("document.pdf", patterns))
        self.assertTrue(matches_ignore_value("path/to/file.PDF", patterns))
        self.assertFalse(matches_ignore_value("document.doc", patterns))
    
    def test_matches_ignore_value_contains_wildcard(self):
        """Test: wildcard z obu stron (*pattern*) - contains."""
        patterns = ["*debug*"]
        
        self.assertTrue(matches_ignore_value("old_debug_temp", patterns))
        self.assertTrue(matches_ignore_value("DEBUG", patterns))
        self.assertTrue(matches_ignore_value("this debug value", patterns))
        self.assertFalse(matches_ignore_value("production", patterns))
    
    def test_matches_ignore_value_multiple_patterns(self):
        """Test: dopasowanie do jednego z wielu wzorców."""
        patterns = ["https", "http*", "*.pdf"]
        
        self.assertTrue(matches_ignore_value("https://example.com", patterns))
        self.assertTrue(matches_ignore_value("http://example.com", patterns))
        self.assertTrue(matches_ignore_value("document.pdf", patterns))
        self.assertFalse(matches_ignore_value("plain text", patterns))
    
    def test_matches_ignore_value_url_examples(self):
        """Test: przykłady z URL-ami (przypadek użytkownika)."""
        patterns = parse_ignore_patterns("https")
        
        # URL-e z https powinny być ignorowane
        self.assertTrue(matches_ignore_value("https://example.com/multiSearch=38858", patterns))
        self.assertTrue(matches_ignore_value("Visit https://example.com for more", patterns))
        self.assertTrue(matches_ignore_value("HTTPS://EXAMPLE.COM", patterns))
        
        # URL-e bez https nie powinny być ignorowane
        self.assertFalse(matches_ignore_value("http://example.com", patterns))
        self.assertFalse(matches_ignore_value("ftp://example.com", patterns))
        self.assertFalse(matches_ignore_value("plain text value", patterns))

    # -------------------- Testy matches_ignore_pattern (zmieniona semantyka) --------------------
    
    def test_matches_ignore_pattern_substring_match_new_semantics(self):
        """Test: nowa semantyka - pattern bez '*' = substring match dla nagłówków."""
        # Nowa semantyka: "test" pasuje do "test_column" (substring match)
        result = matches_ignore_pattern("test_column", ["test"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("my_test_header", ["test"])
        self.assertTrue(result)
        
        # Wildcard patterns pozostają bez zmian
        result = matches_ignore_pattern("test_column", ["test*"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("my_test", ["*test"])
        self.assertTrue(result)
    
    def test_matches_ignore_pattern_backward_compatibility(self):
        """Test: kompatybilność wsteczna - wildcard patterns działają jak wcześniej."""
        # Prefix wildcard
        self.assertTrue(matches_ignore_pattern("temporary", ["temp*"]))
        self.assertFalse(matches_ignore_pattern("notemp", ["temp*"]))
        
        # Suffix wildcard
        self.assertTrue(matches_ignore_pattern("old_temp", ["*temp"]))
        self.assertFalse(matches_ignore_pattern("temporary", ["*temp"]))
        
        # Contains wildcard
        self.assertTrue(matches_ignore_pattern("old_debug_temp", ["*debug*"]))
        self.assertFalse(matches_ignore_pattern("production", ["*debug*"]))
    
    def test_matches_ignore_pattern_exact_match_still_works(self):
        """Test: dokładne dopasowanie nadal działa (jako przypadek szczególny substring)."""
        result = matches_ignore_pattern("test", ["test"])
        self.assertTrue(result)
        
        result = matches_ignore_pattern("TEST", ["test"])
        self.assertTrue(result)

    # -------------------- Testy integracyjne (symulacja przypadków użytkownika) --------------------
    
    def test_user_case_https_in_ignore(self):
        """Test: przypadek użytkownika - 'https' w Ignoruj pomija URL-e."""
        # Parse user input
        ignore_patterns = parse_ignore_patterns("https")
        
        # Symulacja dopasowanych wartości z arkusza
        matched_values = [
            "12345",  # numer zlecenia bez URL
            "https://example.com/multiSearch=38858",  # URL z https
            "67890",  # inny numer
            "Check https://site.com for details",  # tekst z URL
            "http://example.com",  # URL bez https
        ]
        
        # Filtrowanie - tylko wartości NIE pasujące do ignore patterns
        filtered = [v for v in matched_values if not matches_ignore_value(v, ignore_patterns)]
        
        # Oczekiwane wyniki: URL-e z 'https' są odfiltrowane
        expected = [
            "12345",
            "67890",
            "http://example.com",
        ]
        
        self.assertEqual(filtered, expected)
    
    def test_user_case_multiple_ignore_patterns(self):
        """Test: ignorowanie wielu wzorców jednocześnie."""
        # User wpisał: "https, temp*, debug"
        ignore_patterns = parse_ignore_patterns("https, temp*, debug")
        
        matched_values = [
            "regular value",
            "https://example.com",
            "temporary_link",
            "debug_info",
            "production",
        ]
        
        filtered = [v for v in matched_values if not matches_ignore_value(v, ignore_patterns)]
        
        expected = [
            "regular value",
            "production",
        ]
        
        self.assertEqual(filtered, expected)
    
    def test_wildcard_patterns_work_for_values(self):
        """Test: wildcard patterns działają dla wartości komórek."""
        ignore_patterns = parse_ignore_patterns("http*, *.tmp, *debug*")
        
        matched_values = [
            "https://example.com",
            "file.tmp",
            "debug_mode_on",
            "normal_value.txt",
        ]
        
        filtered = [v for v in matched_values if not matches_ignore_value(v, ignore_patterns)]
        
        expected = [
            "normal_value.txt",
        ]
        
        self.assertEqual(filtered, expected)


if __name__ == "__main__":
    # Uruchom testy
    unittest.main(verbosity=2)
