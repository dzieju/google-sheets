"""
test_integration_ignore_values.py
Testy integracyjne dla ignorowania wartości komórek.

Testy symulujące rzeczywiste przypadki użytkownika z arkuszami Google Sheets.
"""

import unittest
from unittest.mock import MagicMock, patch
from sheets_search import (
    search_in_sheet,
    parse_ignore_patterns,
)


class TestIgnoreValuesIntegration(unittest.TestCase):
    """Testy integracyjne ignorowania wartości komórek."""

    def setUp(self):
        """Przygotowanie mocków dla testów."""
        self.mock_drive_service = MagicMock()
        self.mock_sheets_service = MagicMock()
    
    def test_ignore_https_urls_in_search_results(self):
        """
        Test: główny przypadek użytkownika - ignorowanie URL-ów z https.
        
        Użytkownik wpisał "https" w polu Ignoruj i oczekuje, że wyniki
        zawierające URL-e z https:// będą pominięte.
        """
        # Symulacja arkusza z danymi
        sheet_values = [
            ["Numer zlecenia", "Stawka"],  # nagłówek
            ["12345", "100"],  # zwykły numer
            ["https://example.com/multiSearch=38858", "200"],  # URL z https
            ["67890", "150"],  # zwykły numer
            ["https://site.com/order/123", "180"],  # inny URL z https
            ["http://example.com/order/456", "120"],  # URL bez https (http)
        ]
        
        # Parse ignore patterns jak użytkownik wpisałby w GUI
        ignore_patterns = parse_ignore_patterns("https")
        
        # Użyj patch do zasymulowania odpowiedzi API
        with patch.object(self.mock_sheets_service, 'spreadsheets') as mock_spreadsheets:
            mock_spreadsheets.return_value.values.return_value.get.return_value.execute.return_value = {
                'values': sheet_values
            }
            
            # Szukamy wszystkich wartości zawierających cyfry lub "http" (żeby dopasować wszystkie wiersze)
            results = list(search_in_sheet(
                self.mock_drive_service,
                self.mock_sheets_service,
                spreadsheet_id="test_id",
                sheet_name="test_sheet",
                pattern="http",  # szukaj wartości z "http" (dopasuje oba typy URL)
                regex=False,
                case_sensitive=False,
                search_column_name="ALL",  # przeszukuj wszystkie kolumny
                spreadsheet_name="Test Spreadsheet",
                stop_event=None,
                ignore_patterns=ignore_patterns,
                header_row_indices=None,
            ))
        
        # Sprawdź że URL-e z https zostały pominięte
        result_values = [r['searchedValue'] for r in results]
        
        # URL-e z https nie powinny być w wynikach (są ignorowane)
        self.assertNotIn("https://example.com/multiSearch=38858", result_values)
        self.assertNotIn("https://site.com/order/123", result_values)
        
        # URL bez https powinien być w wynikach (nie jest ignorowany)
        self.assertIn("http://example.com/order/456", result_values)
    
    def test_ignore_multiple_patterns_on_values(self):
        """Test: ignorowanie wielu wzorców jednocześnie."""
        sheet_values = [
            ["Data", "Wartość"],
            ["2024-01-01", "normal"],
            ["2024-01-02", "https://example.com"],
            ["2024-01-03", "temp_value"],
            ["2024-01-04", "debug_info"],
            ["2024-01-05", "production"],
        ]
        
        ignore_patterns = parse_ignore_patterns("https, temp*, *debug*")
        
        with patch.object(self.mock_sheets_service, 'spreadsheets') as mock_spreadsheets:
            mock_spreadsheets.return_value.values.return_value.get.return_value.execute.return_value = {
                'values': sheet_values
            }
            
            # Szukamy wartości zawierających "o" (dopasuje "normal" i "production")
            results = list(search_in_sheet(
                self.mock_drive_service,
                self.mock_sheets_service,
                spreadsheet_id="test_id",
                sheet_name="test_sheet",
                pattern="o",  # dopasuje wiele wartości zawierających "o"
                regex=False,
                case_sensitive=False,
                search_column_name="ALL",
                spreadsheet_name="Test Spreadsheet",
                stop_event=None,
                ignore_patterns=ignore_patterns,
                header_row_indices=None,
            ))
        
        result_values = [r['searchedValue'] for r in results]
        
        # Wartości pasujące do ignore patterns nie powinny być w wynikach
        self.assertNotIn("https://example.com", result_values)
        self.assertNotIn("temp_value", result_values)
        self.assertNotIn("debug_info", result_values)
        
        # "normal" i "production" powinny być w wynikach (zawierają 'o' i nie są ignorowane)
        self.assertIn("normal", result_values)
        self.assertIn("production", result_values)
    
    def test_ignore_values_with_header_ignore_combination(self):
        """Test: kombinacja ignorowania nagłówków i wartości."""
        sheet_values = [
            ["Numer", "Temp_Column", "Debug_Data", "Normal"],
            ["12345", "value1", "data1", "normal1"],
            ["https://url.com", "value2", "data2", "normal2"],
            ["67890", "value3", "data3", "normal3"],
        ]
        
        # Ignoruj kolumny zaczynające się na "temp" lub "debug" oraz wartości z "https"
        ignore_patterns = parse_ignore_patterns("temp*, debug*, https")
        
        with patch.object(self.mock_sheets_service, 'spreadsheets') as mock_spreadsheets:
            mock_spreadsheets.return_value.values.return_value.get.return_value.execute.return_value = {
                'values': sheet_values
            }
            
            # Szukamy wartości zawierających cyfry lub "normal"
            results = list(search_in_sheet(
                self.mock_drive_service,
                self.mock_sheets_service,
                spreadsheet_id="test_id",
                sheet_name="test_sheet",
                pattern="normal",  # dopasuje wartości z kolumny Normal
                regex=False,
                case_sensitive=False,
                search_column_name="ALL",
                spreadsheet_name="Test Spreadsheet",
                stop_event=None,
                ignore_patterns=ignore_patterns,
                header_row_indices=None,
            ))
        
        result_values = [r['searchedValue'] for r in results]
        
        # Kolumny "Temp_Column" i "Debug_Data" nie powinny być w ogóle przeszukiwane
        self.assertNotIn("value1", result_values)
        self.assertNotIn("value2", result_values)
        self.assertNotIn("value3", result_values)
        self.assertNotIn("data1", result_values)
        self.assertNotIn("data2", result_values)
        self.assertNotIn("data3", result_values)
        
        # Wartość "https://url.com" z kolumny "Numer" powinna być ignorowana (nie będzie dopasowana przez pattern)
        self.assertNotIn("https://url.com", result_values)
        
        # Wartości z kolumny "Normal" powinny być w wynikach
        self.assertIn("normal1", result_values)
        self.assertIn("normal2", result_values)
        self.assertIn("normal3", result_values)
    
    def test_backward_compatibility_empty_ignore(self):
        """Test: kompatybilność wsteczna - puste ignore działa jak wcześniej."""
        sheet_values = [
            ["Numer", "Wartość"],
            ["https://example.com", "100"],
            ["12345", "200"],
        ]
        
        # Brak ignore patterns
        ignore_patterns = parse_ignore_patterns("")
        
        with patch.object(self.mock_sheets_service, 'spreadsheets') as mock_spreadsheets:
            mock_spreadsheets.return_value.values.return_value.get.return_value.execute.return_value = {
                'values': sheet_values
            }
            
            # Szukamy wartości z "http"
            results = list(search_in_sheet(
                self.mock_drive_service,
                self.mock_sheets_service,
                spreadsheet_id="test_id",
                sheet_name="test_sheet",
                pattern="http",  # dopasuje URL
                regex=False,
                case_sensitive=False,
                search_column_name="ALL",
                spreadsheet_name="Test Spreadsheet",
                stop_event=None,
                ignore_patterns=ignore_patterns,
                header_row_indices=None,
            ))
        
        result_values = [r['searchedValue'] for r in results]
        
        # Bez ignore patterns URL z https powinien być zwrócony
        self.assertIn("https://example.com", result_values)


if __name__ == "__main__":
    unittest.main(verbosity=2)
