"""
test_integration_multi_column.py
Testy integracyjne dla funkcjonalności wyszukiwania wielu kolumn.

Symuluje rzeczywiste scenariusze z wieloma kolumnami o tej samej nazwie.
"""

import unittest
from unittest.mock import MagicMock, patch
from sheets_search import (
    search_in_sheet,
    find_duplicates_in_sheet,
)


class TestMultiColumnIntegration(unittest.TestCase):
    """Testy integracyjne dla wyszukiwania wielu kolumn."""

    def setUp(self):
        """Przygotowanie mocków serwisów."""
        self.mock_drive_service = MagicMock()
        self.mock_sheets_service = MagicMock()

    def test_search_in_sheet_multiple_columns_same_name(self):
        """
        Test: wyszukiwanie w arkuszu z wieloma kolumnami o tej samej nazwie.
        
        Scenariusz:
        - Arkusz ma 3 kolumny: "Zlecenie", "Stawka", "Zlecenie"
        - Wyszukujemy wartość "12345" w kolumnie "Zlecenie"
        - Powinniśmy znaleźć dopasowania z obu kolumn "Zlecenie"
        """
        # Mock danych z arkusza
        mock_values = [
            ["Zlecenie", "Stawka", "Zlecenie", "Inne"],  # Nagłówki
            ["12345", "100", "67890", "x"],              # Wiersz 1
            ["54321", "200", "12345", "y"],              # Wiersz 2
            ["12345", "150", "11111", "z"],              # Wiersz 3
        ]
        
        self.mock_sheets_service.spreadsheets().values().get().execute.return_value = {
            "values": mock_values
        }
        
        # Wykonaj wyszukiwanie
        results = list(search_in_sheet(
            self.mock_drive_service,
            self.mock_sheets_service,
            spreadsheet_id="test_id",
            sheet_name="Sheet1",
            pattern="12345",
            regex=False,
            case_sensitive=False,
            search_column_name="Zlecenie",
            spreadsheet_name="Test Spreadsheet"
        ))
        
        # Weryfikacja wyników
        # Powinniśmy znaleźć 3 dopasowania:
        # - wiersz 1, kolumna A (indeks 0): "12345"
        # - wiersz 2, kolumna C (indeks 2): "12345"
        # - wiersz 3, kolumna A (indeks 0): "12345"
        self.assertEqual(len(results), 3)
        
        # Sprawdź czy wszystkie wyniki mają poprawną strukturę
        for result in results:
            self.assertEqual(result["spreadsheetId"], "test_id")
            self.assertEqual(result["spreadsheetName"], "Test Spreadsheet")
            self.assertEqual(result["sheetName"], "Sheet1")
            self.assertEqual(result["searchedValue"], "12345")
        
        # Sprawdź konkretne komórki
        cells = [r["cell"] for r in results]
        self.assertIn("A2", cells)  # Wiersz 1, kolumna A (12345)
        self.assertIn("C3", cells)  # Wiersz 2, kolumna C (12345)
        self.assertIn("A4", cells)  # Wiersz 3, kolumna A (12345)

    def test_search_in_sheet_case_insensitive_headers(self):
        """
        Test: wyszukiwanie z różnymi wariantami nazw kolumn.
        
        Scenariusz:
        - Arkusz ma kolumny: "zlecenie", "ZLECENIE", "Zlecenie"
        - Wszystkie powinny być traktowane jako ta sama kolumna
        """
        mock_values = [
            ["zlecenie", "Stawka", "ZLECENIE", "Zlecenie"],
            ["AAA", "100", "BBB", "CCC"],
            ["DDD", "200", "EEE", "AAA"],
        ]
        
        self.mock_sheets_service.spreadsheets().values().get().execute.return_value = {
            "values": mock_values
        }
        
        results = list(search_in_sheet(
            self.mock_drive_service,
            self.mock_sheets_service,
            spreadsheet_id="test_id",
            sheet_name="Sheet1",
            pattern="AAA",
            regex=False,
            case_sensitive=False,
            search_column_name="Zlecenie",  # Case-insensitive
            spreadsheet_name="Test Spreadsheet"
        ))
        
        # Powinniśmy znaleźć dopasowania we wszystkich trzech kolumnach "zlecenie"
        self.assertEqual(len(results), 2)  # AAA w kolumnie A wiersz 1 i kolumnie D wiersz 2
        
        cells = [r["cell"] for r in results]
        self.assertIn("A2", cells)  # zlecenie (A) = AAA
        self.assertIn("D3", cells)  # Zlecenie (D) = AAA

    def test_find_duplicates_multiple_columns(self):
        """
        Test: wykrywanie duplikatów w wielu kolumnach o tej samej nazwie.
        
        Scenariusz:
        - Arkusz ma dwie kolumny "Zlecenie" (A i C)
        - Każda kolumna ma swoje duplikaty
        - Duplikaty powinny być raportowane osobno dla każdej kolumny
        """
        mock_values = [
            ["Zlecenie", "Stawka", "Zlecenie"],
            ["12345", "100", "AAA"],
            ["12345", "200", "BBB"],
            ["67890", "150", "AAA"],
            ["12345", "180", "CCC"],
        ]
        
        self.mock_sheets_service.spreadsheets().values().get().execute.return_value = {
            "values": mock_values
        }
        
        results = find_duplicates_in_sheet(
            self.mock_drive_service,
            self.mock_sheets_service,
            spreadsheet_id="test_id",
            sheet_name="Sheet1",
            search_column_name="Zlecenie",
            normalize=True,
            spreadsheet_name="Test Spreadsheet"
        )
        
        # Powinniśmy znaleźć 2 grupy duplikatów:
        # - Kolumna A: "12345" występuje 3 razy (wiersze 2, 3, 5)
        # - Kolumna C: "AAA" występuje 2 razy (wiersze 2, 4)
        self.assertEqual(len(results), 2)
        
        # Sprawdź duplikaty w pierwszej kolumnie (A)
        dup_col_a = [d for d in results if "12345" in d["value"]]
        self.assertEqual(len(dup_col_a), 1)
        self.assertEqual(dup_col_a[0]["count"], 3)
        self.assertEqual(dup_col_a[0]["rows"], [2, 3, 5])
        
        # Sprawdź duplikaty w drugiej kolumnie (C)
        dup_col_c = [d for d in results if "AAA" in d["value"]]
        self.assertEqual(len(dup_col_c), 1)
        self.assertEqual(dup_col_c[0]["count"], 2)
        self.assertEqual(dup_col_c[0]["rows"], [2, 4])

    def test_search_with_whitespace_in_headers(self):
        """
        Test: wyszukiwanie z nagłówkami zawierającymi różne białe znaki.
        
        Scenariusz:
        - Nagłówki: "Numer zlecenia", " Numer zlecenia ", "Numer  zlecenia"
        - Wszystkie powinny być traktowane jako ta sama kolumna
        """
        mock_values = [
            ["Numer zlecenia", "Stawka", " Numer zlecenia ", "Numer  zlecenia"],
            ["111", "100", "222", "333"],
            ["111", "200", "111", "111"],
        ]
        
        self.mock_sheets_service.spreadsheets().values().get().execute.return_value = {
            "values": mock_values
        }
        
        results = list(search_in_sheet(
            self.mock_drive_service,
            self.mock_sheets_service,
            spreadsheet_id="test_id",
            sheet_name="Sheet1",
            pattern="111",
            regex=False,
            case_sensitive=False,
            search_column_name="Numer zlecenia",
            spreadsheet_name="Test Spreadsheet"
        ))
        
        # Powinniśmy znaleźć 4 dopasowania:
        # - Wiersz 1, Kolumna A (index 0): "111"
        # - Wiersz 2, Kolumna A (index 0): "111"
        # - Wiersz 2, Kolumna C (index 2): "111"
        # - Wiersz 2, Kolumna D (index 3): "111"
        self.assertEqual(len(results), 4)

    def test_search_all_mode_with_duplicate_headers(self):
        """
        Test: tryb "ALL" z duplikowanymi nagłówkami.
        
        Scenariusz:
        - Arkusz ma powtarzające się nagłówki
        - Tryb "ALL" powinien przeszukać wszystkie kolumny
        """
        mock_values = [
            ["A", "B", "A", "C"],
            ["test", "x", "y", "test"],
            ["z", "test", "test", "w"],
        ]
        
        self.mock_sheets_service.spreadsheets().values().get().execute.return_value = {
            "values": mock_values
        }
        
        results = list(search_in_sheet(
            self.mock_drive_service,
            self.mock_sheets_service,
            spreadsheet_id="test_id",
            sheet_name="Sheet1",
            pattern="test",
            regex=False,
            case_sensitive=False,
            search_column_name="ALL",
            spreadsheet_name="Test Spreadsheet"
        ))
        
        # Powinniśmy znaleźć 4 dopasowania w różnych miejscach
        self.assertEqual(len(results), 4)


if __name__ == "__main__":
    unittest.main(verbosity=2)
