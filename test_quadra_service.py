"""
test_quadra_service.py
Unit tests for quadra_service module.
"""

import unittest
import tempfile
import os
from unittest.mock import MagicMock, patch
from quadra_service import (
    column_letter_to_index,
    parse_column_identifier,
    read_dbf_column,
    normalize_value_for_comparison,
    values_match,
    search_value_in_sheet_data,
    format_quadra_result_for_table,
    export_quadra_results_to_json,
    export_quadra_results_to_csv,
)


class TestColumnParsing(unittest.TestCase):
    """Tests for column identifier parsing functions."""
    
    def test_column_letter_to_index(self):
        """Test conversion of column letters to indices."""
        self.assertEqual(column_letter_to_index('A'), 0)
        self.assertEqual(column_letter_to_index('B'), 1)
        self.assertEqual(column_letter_to_index('Z'), 25)
        self.assertEqual(column_letter_to_index('AA'), 26)
        self.assertEqual(column_letter_to_index('AB'), 27)
        self.assertEqual(column_letter_to_index('AZ'), 51)
    
    def test_column_letter_case_insensitive(self):
        """Test that column letter parsing is case-insensitive."""
        self.assertEqual(column_letter_to_index('a'), 0)
        self.assertEqual(column_letter_to_index('b'), 1)
        self.assertEqual(column_letter_to_index('aa'), 26)
    
    def test_parse_column_identifier_letters(self):
        """Test parsing column identifiers as letters."""
        self.assertEqual(parse_column_identifier('A'), 0)
        self.assertEqual(parse_column_identifier('B'), 1)
        self.assertEqual(parse_column_identifier('C'), 2)
        self.assertEqual(parse_column_identifier('AA'), 26)
    
    def test_parse_column_identifier_numbers(self):
        """Test parsing column identifiers as 1-based numbers."""
        self.assertEqual(parse_column_identifier('1'), 0)
        self.assertEqual(parse_column_identifier('2'), 1)
        self.assertEqual(parse_column_identifier('10'), 9)
    
    def test_parse_column_identifier_int(self):
        """Test parsing column identifiers as integers."""
        self.assertEqual(parse_column_identifier(1), 0)
        self.assertEqual(parse_column_identifier(2), 1)
        self.assertEqual(parse_column_identifier(10), 9)


class TestValueNormalization(unittest.TestCase):
    """Tests for value normalization and matching functions."""
    
    def test_normalize_value_for_comparison_basic(self):
        """Test basic value normalization."""
        self.assertEqual(normalize_value_for_comparison('Test'), 'test')
        self.assertEqual(normalize_value_for_comparison('  Test  '), 'test')
        self.assertEqual(normalize_value_for_comparison('TEST'), 'test')
    
    def test_normalize_value_for_comparison_numbers(self):
        """Test normalization of numeric values."""
        # The normalize_number_string should handle these
        result = normalize_value_for_comparison('12345')
        self.assertEqual(result, '12345')
        
        result = normalize_value_for_comparison(12345)
        self.assertEqual(result, '12345')
    
    def test_normalize_value_for_comparison_none(self):
        """Test normalization of None values."""
        self.assertEqual(normalize_value_for_comparison(None), '')
    
    def test_values_match_exact(self):
        """Test exact value matching."""
        self.assertTrue(values_match('Test', 'test', mode='exact'))
        self.assertTrue(values_match('  Test  ', 'test', mode='exact'))
        self.assertTrue(values_match('12345', '12345', mode='exact'))
        self.assertFalse(values_match('Test', 'Test2', mode='exact'))
    
    def test_values_match_substring(self):
        """Test substring value matching."""
        self.assertTrue(values_match('Test', 'Testing', mode='substring'))
        self.assertTrue(values_match('Testing', 'Test', mode='substring'))
        self.assertTrue(values_match('123', '12345', mode='substring'))
        self.assertFalse(values_match('abc', 'xyz', mode='substring'))
    
    def test_values_match_none(self):
        """Test matching with None values."""
        self.assertTrue(values_match(None, None, mode='exact'))
        self.assertFalse(values_match('Test', None, mode='exact'))
        self.assertFalse(values_match(None, 'Test', mode='exact'))


class TestSearchValueInSheetData(unittest.TestCase):
    """Tests for searching values in sheet data."""
    
    def test_search_value_found(self):
        """Test finding a value in sheet data."""
        sheet_data = [
            ['ID', 'Order', 'Description'],
            [1, '12345', 'Test 1'],
            [2, '67890', 'Test 2'],
            [3, 'ABC-001', 'Test 3'],
        ]
        
        result = search_value_in_sheet_data(
            target_value='12345',
            sheet_values=sheet_data,
            sheet_name='Sheet1',
            mode='exact'
        )
        
        self.assertIsNotNone(result)
        self.assertEqual(result['sheetName'], 'Sheet1')
        self.assertEqual(result['columnIndex'], 1)
        self.assertEqual(result['rowIndex'], 1)
        self.assertEqual(result['value'], '12345')
    
    def test_search_value_not_found(self):
        """Test when value is not found in sheet data."""
        sheet_data = [
            ['ID', 'Order', 'Description'],
            [1, '12345', 'Test 1'],
            [2, '67890', 'Test 2'],
        ]
        
        result = search_value_in_sheet_data(
            target_value='99999',
            sheet_values=sheet_data,
            sheet_name='Sheet1',
            mode='exact'
        )
        
        self.assertIsNone(result)
    
    def test_search_value_empty_sheet(self):
        """Test searching in empty sheet data."""
        result = search_value_in_sheet_data(
            target_value='12345',
            sheet_values=[],
            sheet_name='Sheet1',
            mode='exact'
        )
        
        self.assertIsNone(result)
    
    def test_search_value_in_specific_columns(self):
        """Test searching only in specific columns."""
        sheet_data = [
            ['ID', 'Order', 'Description'],
            [1, '12345', 'Test with 12345'],
            [2, '67890', 'Test 2'],
        ]
        
        # Search only in 'Order' column
        result = search_value_in_sheet_data(
            target_value='12345',
            sheet_values=sheet_data,
            sheet_name='Sheet1',
            mode='exact',
            column_names=['Order']
        )
        
        self.assertIsNotNone(result)
        self.assertEqual(result['columnIndex'], 1)
        self.assertEqual(result['columnName'], 'Order')


class TestResultFormatting(unittest.TestCase):
    """Tests for result formatting functions."""
    
    def test_format_quadra_result_for_table_found(self):
        """Test formatting a found result for table display."""
        result = {
            'dbfValue': '12345',
            'found': True,
            'sheetName': 'Sheet1',
            'columnName': 'Order',
            'rowIndex': 5,
            'notes': 'Found in Sheet1 at B6'
        }
        
        table_row = format_quadra_result_for_table(result)
        
        self.assertEqual(table_row[0], '12345')  # DBF value
        self.assertEqual(table_row[1], 'Found')  # Status
        self.assertEqual(table_row[2], 'Sheet1')  # Sheet name
        self.assertEqual(table_row[3], 'Order')  # Column name
        self.assertEqual(table_row[4], '6')  # Row (1-based)
        self.assertEqual(table_row[5], 'Found in Sheet1 at B6')  # Notes
    
    def test_format_quadra_result_for_table_missing(self):
        """Test formatting a missing result for table display."""
        result = {
            'dbfValue': '99999',
            'found': False,
            'sheetName': None,
            'columnName': None,
            'rowIndex': None,
            'notes': 'Missing'
        }
        
        table_row = format_quadra_result_for_table(result)
        
        self.assertEqual(table_row[0], '99999')  # DBF value
        self.assertEqual(table_row[1], 'Missing')  # Status
        self.assertEqual(table_row[2], '')  # Sheet name
        self.assertEqual(table_row[3], '')  # Column name
        self.assertEqual(table_row[4], '')  # Row
        self.assertEqual(table_row[5], 'Missing')  # Notes
    
    def test_export_quadra_results_to_json(self):
        """Test JSON export formatting."""
        results = [
            {
                'dbfValue': '12345',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Order',
                'columnIndex': 1,
                'rowIndex': 5,
                'matchedValue': '12345',
                'notes': 'Found'
            },
            {
                'dbfValue': '99999',
                'found': False,
                'sheetName': None,
                'columnName': None,
                'columnIndex': None,
                'rowIndex': None,
                'matchedValue': None,
                'notes': 'Missing'
            }
        ]
        
        json_results = export_quadra_results_to_json(results)
        
        self.assertEqual(len(json_results), 2)
        self.assertEqual(json_results[0]['status'], 'Found')
        self.assertEqual(json_results[1]['status'], 'Missing')
    
    def test_export_quadra_results_to_csv(self):
        """Test CSV export formatting."""
        results = [
            {
                'dbfValue': '12345',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Order',
                'columnIndex': 1,
                'rowIndex': 5,
                'matchedValue': '12345',
                'notes': 'Found'
            }
        ]
        
        csv_output = export_quadra_results_to_csv(results)
        
        self.assertIn('DBF_Value', csv_output)
        self.assertIn('Status', csv_output)
        self.assertIn('12345', csv_output)
        self.assertIn('Found', csv_output)


class TestDBFReading(unittest.TestCase):
    """Tests for DBF reading functions."""
    
    def setUp(self):
        """Create a temporary DBF file for testing."""
        import dbf
        
        # Create temporary DBF file
        self.temp_dir = tempfile.mkdtemp()
        self.dbf_path = os.path.join(self.temp_dir, 'test.dbf')
        
        table = dbf.Table(self.dbf_path, 'id N(10,0); order_num C(20); desc C(50)')
        table.open(mode=dbf.READ_WRITE)
        table.append((1, '12345', 'Test 1'))
        table.append((2, '67890', 'Test 2'))
        table.append((3, 'ABC-001', 'Test 3'))
        table.close()
    
    def tearDown(self):
        """Clean up temporary files."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_read_dbf_column_by_letter(self):
        """Test reading DBF column by letter identifier."""
        values = read_dbf_column(self.dbf_path, 'B')
        self.assertEqual(len(values), 3)
        self.assertEqual(values[0], '12345')
        self.assertEqual(values[1], '67890')
        self.assertEqual(values[2], 'ABC-001')
    
    def test_read_dbf_column_by_number(self):
        """Test reading DBF column by numeric identifier."""
        values = read_dbf_column(self.dbf_path, '2')
        self.assertEqual(len(values), 3)
        self.assertEqual(values[0], '12345')
    
    def test_read_dbf_column_invalid_file(self):
        """Test reading from non-existent DBF file."""
        with self.assertRaises(FileNotFoundError):
            read_dbf_column('/nonexistent/file.dbf', 'B')
    
    def test_read_dbf_column_invalid_column(self):
        """Test reading from invalid column."""
        with self.assertRaises(ValueError):
            read_dbf_column(self.dbf_path, 'ZZ')


if __name__ == '__main__':
    unittest.main()
