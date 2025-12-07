"""
test_quadra_service.py
Unit tests for quadra_service module.

Test Dependencies:
- dbf (python-dbf): Used to CREATE test DBF files
  Install with: pip install dbf
  Note: Production code uses 'dbfread' for reading, which is lighter and read-only.
"""

import unittest
import tempfile
import os
import shutil
import dbf
from unittest.mock import MagicMock, patch
from quadra_service import (
    column_letter_to_index,
    parse_column_identifier,
    read_dbf_column,
    detect_dbf_field_name,
    map_dbf_record_to_result,
    read_dbf_records_with_extra_fields,
    normalize_value_for_comparison,
    values_match,
    search_value_in_sheet_data,
    search_dbf_values_in_sheets,
    format_quadra_result_for_table,
    map_column_names,
    export_quadra_results_to_json,
    export_quadra_results_to_csv,
    write_quadra_results_to_sheet,
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
            'notes': 'Found in Sheet1 at B6',
            'stawka': '',
            'czesci': '',
            'platnik': ''
        }
        
        table_row = format_quadra_result_for_table(result)
        
        # New format: [Arkusz, Płatnik, Numer z DBF, Stawka, Czesci, Status, Kolumna, Wiersz, Uwagi]
        self.assertEqual(table_row[0], 'Sheet1')  # Arkusz (Sheet name)
        self.assertEqual(table_row[1], '')  # Płatnik
        self.assertEqual(table_row[2], '12345')  # Numer z DBF (DBF value)
        self.assertEqual(table_row[3], '')  # Stawka
        self.assertEqual(table_row[4], '')  # Czesci
        self.assertEqual(table_row[5], 'Found')  # Status
        self.assertEqual(table_row[6], 'Order')  # Kolumna (Column name)
        self.assertEqual(table_row[7], '6')  # Wiersz (Row, 1-based)
        self.assertEqual(table_row[8], 'Found in Sheet1 at B6')  # Uwagi (Notes)
    
    def test_format_quadra_result_for_table_missing(self):
        """Test formatting a missing result for table display."""
        result = {
            'dbfValue': '99999',
            'found': False,
            'sheetName': None,
            'columnName': None,
            'rowIndex': None,
            'notes': 'Missing',
            'stawka': '',
            'czesci': '',
            'platnik': ''
        }
        
        table_row = format_quadra_result_for_table(result)
        
        # New format: [Arkusz, Płatnik, Numer z DBF, Stawka, Czesci, Status, Kolumna, Wiersz, Uwagi]
        self.assertEqual(table_row[0], '')  # Arkusz (Sheet name)
        self.assertEqual(table_row[1], '')  # Płatnik
        self.assertEqual(table_row[2], '99999')  # Numer z DBF (DBF value)
        self.assertEqual(table_row[3], '')  # Stawka
        self.assertEqual(table_row[4], '')  # Czesci
        self.assertEqual(table_row[5], 'Missing')  # Status
        self.assertEqual(table_row[6], '')  # Kolumna (Column name)
        self.assertEqual(table_row[7], '')  # Wiersz (Row)
        self.assertEqual(table_row[8], 'Missing')  # Uwagi (Notes)
    
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
                'notes': 'Found',
                'stawka': '',
                'czesci': ''
            },
            {
                'dbfValue': '99999',
                'found': False,
                'sheetName': None,
                'columnName': None,
                'columnIndex': None,
                'rowIndex': None,
                'matchedValue': None,
                'notes': 'Missing',
                'stawka': '',
                'czesci': ''
            }
        ]
        
        json_results = export_quadra_results_to_json(results)
        
        self.assertEqual(len(json_results), 2)
        self.assertEqual(json_results[0]['status'], 'Found')
        self.assertEqual(json_results[1]['status'], 'Missing')
        # Check that stawka and czesci are included
        self.assertIn('stawka', json_results[0])
        self.assertIn('czesci', json_results[0])
    
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
                'notes': 'Found',
                'stawka': '',
                'czesci': ''
            }
        ]
        
        csv_output = export_quadra_results_to_csv(results)
        
        self.assertIn('DBF_Value', csv_output)
        self.assertIn('Status', csv_output)
        self.assertIn('12345', csv_output)
        self.assertIn('Found', csv_output)


class TestDBFReading(unittest.TestCase):
    """Tests for DBF reading functions.
    
    Note: These tests use the 'dbf' library (python-dbf) to CREATE test DBF files,
    while the production code uses 'dbfread' to READ them. This is intentional:
    - dbfread: Fast, read-only, no dependencies - perfect for production
    - dbf: Can write DBF files - needed for creating test fixtures
    Both libraries are compatible with the same DBF file format.
    """
    
    def setUp(self):
        """Create a temporary DBF file for testing."""
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


class TestDBFFieldDetection(unittest.TestCase):
    """Tests for DBF field name detection functions."""
    
    def test_detect_dbf_field_name_exact_match(self):
        """Test detecting exact field name match."""
        field_names = ['NUMER', 'STAWKA', 'DATA']
        possible_names = ['STAWKA', 'STAW', 'RATE']
        
        result = detect_dbf_field_name(field_names, possible_names)
        self.assertEqual(result, 'STAWKA')
    
    def test_detect_dbf_field_name_alternative(self):
        """Test detecting alternative field name."""
        field_names = ['NUMER', 'RATE', 'DATA']
        possible_names = ['STAWKA', 'STAW', 'RATE']
        
        result = detect_dbf_field_name(field_names, possible_names)
        self.assertEqual(result, 'RATE')
    
    def test_detect_dbf_field_name_case_insensitive(self):
        """Test case-insensitive field name detection."""
        field_names = ['numer', 'stawka', 'data']
        possible_names = ['STAWKA', 'STAW', 'RATE']
        
        result = detect_dbf_field_name(field_names, possible_names)
        self.assertEqual(result, 'stawka')
    
    def test_detect_dbf_field_name_not_found(self):
        """Test when field name is not found."""
        field_names = ['NUMER', 'DATA', 'UWAGI']
        possible_names = ['STAWKA', 'STAW', 'RATE']
        
        result = detect_dbf_field_name(field_names, possible_names)
        self.assertIsNone(result)
    
    def test_detect_dbf_field_name_empty_lists(self):
        """Test with empty input lists."""
        self.assertIsNone(detect_dbf_field_name([], ['STAWKA']))
        self.assertIsNone(detect_dbf_field_name(['NUMER'], []))
        self.assertIsNone(detect_dbf_field_name([], []))


class TestMapDBFRecord(unittest.TestCase):
    """Tests for mapping DBF records to results."""
    
    def test_map_dbf_record_with_all_fields(self):
        """Test mapping record with numer_dbf, stawka and czesci fields."""
        record = {
            'NUMER': '12345',
            'STAWKA': '150.00',
            'CZESCI': 'ABC',
            'DATA': '2025-01-01'
        }
        field_names = ['NUMER', 'STAWKA', 'CZESCI', 'DATA']
        
        result = map_dbf_record_to_result(record, field_names)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '150.00')
        self.assertEqual(result['czesci'], 'ABC')
        self.assertEqual(result['platnik'], '')  # Platnik not mapped, should be empty
    
    def test_map_dbf_record_with_alternative_names(self):
        """Test mapping record with alternative field names."""
        record = {
            'ORDER': '12345',
            'RATE': '150.00',
            'PARTS': 'ABC'
        }
        field_names = ['ORDER', 'RATE', 'PARTS']
        
        result = map_dbf_record_to_result(record, field_names)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '150.00')
        self.assertEqual(result['czesci'], 'ABC')
        self.assertEqual(result['platnik'], '')  # Platnik not mapped, should be empty
    
    def test_map_dbf_record_missing_fields(self):
        """Test mapping record with missing fields."""
        record = {
            'NUMER': '12345',
            'DATA': '2025-01-01'
        }
        field_names = ['NUMER', 'DATA']
        
        result = map_dbf_record_to_result(record, field_names)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '')
        self.assertEqual(result['czesci'], '')
        self.assertEqual(result['platnik'], '')  # Platnik not present
    
    def test_map_dbf_record_with_none_values(self):
        """Test mapping record with None values."""
        record = {
            'NUMER': '12345',
            'STAWKA': None,
            'CZESCI': None
        }
        field_names = ['NUMER', 'STAWKA', 'CZESCI']
        
        result = map_dbf_record_to_result(record, field_names)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '')
        self.assertEqual(result['czesci'], '')
        self.assertEqual(result['platnik'], '')  # Platnik not present
    
    def test_map_dbf_record_with_user_mapping(self):
        """Test mapping record with user-provided mapping."""
        record = {
            'ORDER_NUM': '12345',
            'PRICE': '150.00',
            'COMPONENTS': 'ABC',
            'DATA': '2025-01-01'
        }
        field_names = ['ORDER_NUM', 'PRICE', 'COMPONENTS', 'DATA']
        
        # User-provided mapping
        mapping = {
            'numer_dbf': 'ORDER_NUM',
            'stawka': 'PRICE',
            'czesci': 'COMPONENTS'
        }
        
        result = map_dbf_record_to_result(record, field_names, mapping)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '150.00')
        self.assertEqual(result['czesci'], 'ABC')
        self.assertEqual(result['platnik'], '')  # Platnik not mapped
    
    def test_map_dbf_record_with_partial_mapping(self):
        """Test mapping record with partial user mapping (some fields auto-detected)."""
        record = {
            'ORDER_NUM': '12345',
            'STAWKA': '150.00',
            'CZESCI': 'ABC'
        }
        field_names = ['ORDER_NUM', 'STAWKA', 'CZESCI']
        
        # User only maps numer_dbf, others auto-detected
        mapping = {
            'numer_dbf': 'ORDER_NUM'
        }
        
        result = map_dbf_record_to_result(record, field_names, mapping)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '150.00')
        self.assertEqual(result['czesci'], 'ABC')
        self.assertEqual(result['platnik'], '')  # Platnik not present
    
    def test_map_dbf_record_with_cena_alternative(self):
        """Test mapping record with CENA as alternative for stawka."""
        record = {
            'NUMER': '12345',
            'CENA': '150.00',
            'CZESCI': 'ABC'
        }
        field_names = ['NUMER', 'CENA', 'CZESCI']
        
        result = map_dbf_record_to_result(record, field_names)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '150.00')
        self.assertEqual(result['czesci'], 'ABC')
        self.assertEqual(result['platnik'], '')  # Platnik not present
    
    def test_map_dbf_record_with_czesc_alternative(self):
        """Test mapping record with CZESC/PART as alternative for czesci."""
        record = {
            'NUMER': '12345',
            'STAWKA': '150.00',
            'PART': 'ABC'
        }
        field_names = ['NUMER', 'STAWKA', 'PART']
        
        result = map_dbf_record_to_result(record, field_names)
        
        self.assertEqual(result['numer_dbf'], '12345')
        self.assertEqual(result['stawka'], '150.00')
        self.assertEqual(result['czesci'], 'ABC')
        self.assertEqual(result['platnik'], '')  # Platnik not present


class TestReadDBFRecordsWithExtraFields(unittest.TestCase):
    """Tests for reading DBF records with extra fields."""
    
    def setUp(self):
        """Set up test DBF file with extra fields."""
        self.temp_dir = tempfile.mkdtemp()
        self.dbf_path = os.path.join(self.temp_dir, 'test_extra.dbf')
        
        # Create DBF with ID, ORDER, STAWKA, CZESCI columns
        table = dbf.Table(self.dbf_path, 'ID N(5,0); ORDER C(20); STAWKA C(10); CZESCI C(10)')
        table.open(mode=dbf.READ_WRITE)
        table.append((1, '12345', '150.00', 'ABC'))
        table.append((2, '67890', '200.50', 'XYZ'))
        table.append((3, 'ABC-001', '100.00', 'DEF'))
        table.close()
    
    def tearDown(self):
        """Clean up test files."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_read_records_with_extra_fields(self):
        """Test reading DBF records with stawka and czesci fields."""
        records = read_dbf_records_with_extra_fields(self.dbf_path, 'B')
        
        self.assertEqual(len(records), 3)
        
        # Check first record
        self.assertEqual(records[0]['value'], '12345')
        self.assertEqual(records[0]['stawka'], '150.00')
        self.assertEqual(records[0]['czesci'], 'ABC')
        self.assertEqual(records[0]['platnik'], '')  # Platnik not present in test DBF
        
        # Check second record
        self.assertEqual(records[1]['value'], '67890')
        self.assertEqual(records[1]['stawka'], '200.50')
        self.assertEqual(records[1]['czesci'], 'XYZ')
        self.assertEqual(records[1]['platnik'], '')  # Platnik not present in test DBF
    
    def test_read_records_missing_extra_fields(self):
        """Test reading DBF records when extra fields are missing."""
        # Create DBF without STAWKA and CZESCI
        dbf_path2 = os.path.join(self.temp_dir, 'test_no_extra.dbf')
        table = dbf.Table(dbf_path2, 'ID N(5,0); ORDER C(20)')
        table.open(mode=dbf.READ_WRITE)
        table.append((1, '12345'))
        table.close()
        
        records = read_dbf_records_with_extra_fields(dbf_path2, 'B')
        
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]['value'], '12345')
        self.assertEqual(records[0]['stawka'], '')
        self.assertEqual(records[0]['czesci'], '')
        self.assertEqual(records[0]['platnik'], '')  # Platnik not present


class TestSearchDBFValuesWithExtraFields(unittest.TestCase):
    """Tests for searching DBF values with extra fields."""
    
    def test_search_with_simple_values(self):
        """Test search with simple values (backward compatibility)."""
        # Mock services
        mock_drive = MagicMock()
        mock_sheets = MagicMock()
        
        # Mock spreadsheet metadata
        mock_sheets.spreadsheets().get().execute.return_value = {
            'sheets': [{'properties': {'title': 'Sheet1'}}]
        }
        
        # Mock sheet data
        mock_sheets.spreadsheets().values().get().execute.return_value = {
            'values': [
                ['ID', 'Order', 'Description'],
                [1, '12345', 'Test 1'],
                [2, '67890', 'Test 2'],
            ]
        }
        
        # Search with simple values
        results = search_dbf_values_in_sheets(
            drive_service=mock_drive,
            sheets_service=mock_sheets,
            dbf_values=['12345', '99999'],
            spreadsheet_id='test_id',
            mode='exact'
        )
        
        self.assertEqual(len(results), 2)
        
        # First value found
        self.assertTrue(results[0]['found'])
        self.assertEqual(results[0]['dbfValue'], '12345')
        self.assertEqual(results[0]['stawka'], '')
        self.assertEqual(results[0]['czesci'], '')
        self.assertEqual(results[0]['platnik'], '')  # Platnik not in simple values
        
        # Second value not found
        self.assertFalse(results[1]['found'])
        self.assertEqual(results[1]['dbfValue'], '99999')
        self.assertEqual(results[1]['stawka'], '')
        self.assertEqual(results[1]['czesci'], '')
        self.assertEqual(results[1]['platnik'], '')  # Platnik not in simple values
    
    def test_search_with_record_dicts(self):
        """Test search with record dictionaries containing extra fields."""
        # Mock services
        mock_drive = MagicMock()
        mock_sheets = MagicMock()
        
        # Mock spreadsheet metadata
        mock_sheets.spreadsheets().get().execute.return_value = {
            'sheets': [{'properties': {'title': 'Sheet1'}}]
        }
        
        # Mock sheet data
        mock_sheets.spreadsheets().values().get().execute.return_value = {
            'values': [
                ['ID', 'Order', 'Description'],
                [1, '12345', 'Test 1'],
            ]
        }
        
        # Search with record dicts
        dbf_records = [
            {'value': '12345', 'stawka': '150.00', 'czesci': 'ABC', 'platnik': 'Company A'},
            {'value': '99999', 'stawka': '200.50', 'czesci': 'XYZ', 'platnik': 'Company B'}
        ]
        
        results = search_dbf_values_in_sheets(
            drive_service=mock_drive,
            sheets_service=mock_sheets,
            dbf_values=dbf_records,
            spreadsheet_id='test_id',
            mode='exact'
        )
        
        self.assertEqual(len(results), 2)
        
        # First record found with extra fields
        self.assertTrue(results[0]['found'])
        self.assertEqual(results[0]['dbfValue'], '12345')
        self.assertEqual(results[0]['stawka'], '150.00')
        self.assertEqual(results[0]['czesci'], 'ABC')
        self.assertEqual(results[0]['platnik'], 'Company A')  # Platnik preserved
        
        # Second record not found with extra fields preserved
        self.assertFalse(results[1]['found'])
        self.assertEqual(results[1]['dbfValue'], '99999')
        self.assertEqual(results[1]['stawka'], '200.50')
        self.assertEqual(results[1]['czesci'], 'XYZ')
        self.assertEqual(results[1]['platnik'], 'Company B')  # Platnik preserved


class TestFormatQuadraResultWithExtraFields(unittest.TestCase):
    """Tests for formatting Quadra results with extra fields."""
    
    def test_format_result_for_table(self):
        """Test formatting result for GUI table with extra fields."""
        result = {
            'dbfValue': '12345',
            'found': True,
            'sheetName': 'Sheet1',
            'columnName': 'Order',
            'rowIndex': 5,
            'notes': 'Found in Sheet1 at B6',
            'stawka': '150.00',
            'czesci': 'ABC',
            'platnik': 'Company XYZ'
        }
        
        row = format_quadra_result_for_table(result)
        
        # Expected format: [Arkusz, Płatnik, Numer z DBF, Stawka, Czesci, Status, Kolumna, Wiersz, Uwagi]
        self.assertEqual(len(row), 9)
        self.assertEqual(row[0], 'Sheet1')  # Arkusz
        self.assertEqual(row[1], 'Company XYZ')  # Płatnik
        self.assertEqual(row[2], '12345')  # Numer z DBF
        self.assertEqual(row[3], '150.00')  # Stawka
        self.assertEqual(row[4], 'ABC')  # Czesci
        self.assertEqual(row[5], 'Found')  # Status
        self.assertEqual(row[6], 'Order')  # Kolumna
        self.assertEqual(row[7], '6')  # Wiersz
        self.assertEqual(row[8], 'Found in Sheet1 at B6')  # Uwagi
    
    def test_format_result_missing(self):
        """Test formatting missing result with extra fields."""
        result = {
            'dbfValue': '99999',
            'found': False,
            'sheetName': None,
            'columnName': None,
            'rowIndex': None,
            'notes': 'Missing',
            'stawka': '200.50',
            'czesci': 'XYZ',
            'platnik': 'Company ABC'
        }
        
        row = format_quadra_result_for_table(result)
        
        # Expected format: [Arkusz, Płatnik, Numer z DBF, Stawka, Czesci, Status, Kolumna, Wiersz, Uwagi]
        self.assertEqual(len(row), 9)
        self.assertEqual(row[0], '')  # Arkusz
        self.assertEqual(row[1], 'Company ABC')  # Płatnik
        self.assertEqual(row[2], '99999')  # Numer z DBF
        self.assertEqual(row[3], '200.50')  # Stawka
        self.assertEqual(row[4], 'XYZ')  # Czesci
        self.assertEqual(row[5], 'Missing')  # Status
        self.assertEqual(row[6], '')  # Kolumna
        self.assertEqual(row[7], '')  # Wiersz
        self.assertEqual(row[8], 'Missing')  # Uwagi


class TestExportQuadraResultsWithExtraFields(unittest.TestCase):
    """Tests for exporting Quadra results with extra fields."""
    
    def test_export_to_json(self):
        """Test JSON export with extra fields."""
        results = [
            {
                'dbfValue': '12345',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Order',
                'columnIndex': 1,
                'rowIndex': 5,
                'matchedValue': '12345',
                'notes': 'Found',
                'stawka': '150.00',
                'czesci': 'ABC'
            }
        ]
        
        export_data = export_quadra_results_to_json(results)
        
        self.assertEqual(len(export_data), 1)
        self.assertEqual(export_data[0]['dbfValue'], '12345')
        self.assertEqual(export_data[0]['stawka'], '150.00')
        self.assertEqual(export_data[0]['czesci'], 'ABC')
        self.assertEqual(export_data[0]['status'], 'Found')
    
    def test_export_to_csv(self):
        """Test CSV export with extra fields."""
        results = [
            {
                'dbfValue': '12345',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Order',
                'columnIndex': 1,
                'rowIndex': 5,
                'matchedValue': '12345',
                'notes': 'Found',
                'stawka': '150.00',
                'czesci': 'ABC'
            }
        ]
        
        csv_data = export_quadra_results_to_csv(results)
        
        lines = csv_data.strip().split('\n')
        self.assertEqual(len(lines), 2)  # Header + 1 data row
        
        # Check header
        header = lines[0]
        self.assertIn('Stawka', header)
        self.assertIn('Czesci', header)
        
        # Check data row
        data_row = lines[1]
        self.assertIn('12345', data_row)
        self.assertIn('150.00', data_row)
        self.assertIn('ABC', data_row)


class TestWriteQuadraResultsToSheet(unittest.TestCase):
    """Tests for writing Quadra results to Google Sheets."""
    
    def test_write_results_to_sheet(self):
        """Test writing results to columns I and J."""
        # Mock services
        mock_sheets = MagicMock()
        
        # Mock spreadsheet metadata
        mock_sheets.spreadsheets().get().execute.return_value = {
            'sheets': [{'properties': {'sheetId': 123, 'title': 'Sheet1'}}]
        }
        
        # Mock update response
        mock_update_result = MagicMock()
        mock_update_result.execute.return_value = {
            'updatedCells': 6  # Header + 2 data rows = 3 rows * 2 columns
        }
        mock_sheets.spreadsheets().values().update.return_value = mock_update_result
        
        results = [
            {'dbfValue': '12345', 'stawka': '150.00', 'czesci': 'ABC', 'found': True},
            {'dbfValue': '67890', 'stawka': '200.50', 'czesci': 'XYZ', 'found': False}
        ]
        
        write_quadra_results_to_sheet(
            sheets_service=mock_sheets,
            spreadsheet_id='test_id',
            sheet_name='Sheet1',
            results=results,
            start_row=1
        )
        
        # Verify update was called with correct parameters
        mock_sheets.spreadsheets().values().update.assert_called_with(
            spreadsheetId='test_id',
            range='Sheet1!I1:J3',
            valueInputOption='RAW',
            body={
                'values': [
                    ['Stawka', 'Czesci'],
                    ['150.00', 'ABC'],
                    ['200.50', 'XYZ']
                ]
            }
        )
    
    def test_write_results_empty_fields(self):
        """Test writing results with empty extra fields."""
        mock_sheets = MagicMock()
        
        mock_sheets.spreadsheets().get().execute.return_value = {
            'sheets': [{'properties': {'sheetId': 123, 'title': 'Sheet1'}}]
        }
        
        mock_update_result = MagicMock()
        mock_update_result.execute.return_value = {
            'updatedCells': 4
        }
        mock_sheets.spreadsheets().values().update.return_value = mock_update_result
        
        results = [
            {'dbfValue': '12345', 'stawka': '', 'czesci': '', 'found': True}
        ]
        
        write_quadra_results_to_sheet(
            sheets_service=mock_sheets,
            spreadsheet_id='test_id',
            sheet_name='Sheet1',
            results=results,
            start_row=1
        )
        
        # Verify update was called with correct empty values
        mock_sheets.spreadsheets().values().update.assert_called_with(
            spreadsheetId='test_id',
            range='Sheet1!I1:J2',
            valueInputOption='RAW',
            body={
                'values': [
                    ['Stawka', 'Czesci'],
                    ['', '']
                ]
            }
        )
    
    def test_write_results_sheet_not_found(self):
        """Test error handling when sheet is not found."""
        mock_sheets = MagicMock()
        
        mock_sheets.spreadsheets().get().execute.return_value = {
            'sheets': [{'properties': {'sheetId': 123, 'title': 'OtherSheet'}}]
        }
        
        results = [{'dbfValue': '12345', 'stawka': '150.00', 'czesci': 'ABC'}]
        
        with self.assertRaises(ValueError) as cm:
            write_quadra_results_to_sheet(
                sheets_service=mock_sheets,
                spreadsheet_id='test_id',
                sheet_name='Sheet1',
                results=results
            )
        
        self.assertIn("Sheet 'Sheet1' not found", str(cm.exception))


class TestMapColumnNames(unittest.TestCase):
    """Tests for map_column_names function."""
    
    def test_map_with_dict(self):
        """Test mapping with dictionary."""
        from quadra_service import map_column_names
        
        original = ['dbfValue', 'stawka', 'status']
        mapping = {'dbfValue': 'Numer', 'stawka': 'Kwota'}
        result = map_column_names(original, mapping)
        
        self.assertEqual(result, ['Numer', 'Kwota', 'status'])
    
    def test_map_with_list(self):
        """Test mapping with list."""
        from quadra_service import map_column_names
        
        original = ['dbfValue', 'stawka', 'status']
        mapping = ['Numer', 'Kwota', 'Stan']
        result = map_column_names(original, mapping)
        
        self.assertEqual(result, ['Numer', 'Kwota', 'Stan'])
    
    def test_map_with_list_fewer_elements(self):
        """Test mapping with list that has fewer elements than original."""
        from quadra_service import map_column_names
        
        original = ['dbfValue', 'stawka', 'status']
        mapping = ['Numer', 'Kwota']  # Missing third element
        result = map_column_names(original, mapping)
        
        # Should use original for missing elements
        self.assertEqual(result, ['Numer', 'Kwota', 'status'])
    
    def test_map_with_none(self):
        """Test mapping with None - should return original."""
        from quadra_service import map_column_names
        
        original = ['dbfValue', 'stawka', 'status']
        result = map_column_names(original, None)
        
        self.assertEqual(result, original)
    
    def test_map_with_empty_dict(self):
        """Test mapping with empty dict - should return original."""
        from quadra_service import map_column_names
        
        original = ['dbfValue', 'stawka', 'status']
        result = map_column_names(original, {})
        
        self.assertEqual(result, original)


class TestExportWithColumnNames(unittest.TestCase):
    """Tests for export functions with custom column names."""
    
    def test_export_csv_with_dict_column_names(self):
        """Test CSV export with dictionary column name mapping."""
        results = [
            {
                'dbfValue': '12345',
                'stawka': '150.00',
                'czesci': 'ABC',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Numer',
                'columnIndex': 0,
                'rowIndex': 1,
                'matchedValue': '12345',
                'notes': 'Test'
            }
        ]
        
        column_names = {
            'DBF_Value': 'Wartość DBF',
            'Stawka': 'Cena',
            'Status': 'Stan'
        }
        
        csv_output = export_quadra_results_to_csv(results, column_names)
        lines = csv_output.strip().split('\n')
        
        # Check header
        header = lines[0]
        self.assertIn('Wartość DBF', header)
        self.assertIn('Cena', header)
        self.assertIn('Stan', header)
    
    def test_export_csv_with_list_column_names(self):
        """Test CSV export with list column name mapping."""
        results = [
            {
                'dbfValue': '12345',
                'stawka': '150.00',
                'czesci': 'ABC',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Numer',
                'columnIndex': 0,
                'rowIndex': 1,
                'matchedValue': '12345',
                'notes': 'Test'
            }
        ]
        
        column_names = ['Wartość', 'Cena', 'Stan', 'Arkusz', 'Kolumna', 
                       'Indeks kolumny', 'Indeks wiersza', 'Znaleziona wartość', 
                       'Części', 'Notatki']
        
        csv_output = export_quadra_results_to_csv(results, column_names)
        lines = csv_output.strip().split('\n')
        
        # Check header
        header = lines[0]
        self.assertIn('Wartość', header)
        self.assertIn('Cena', header)
        self.assertIn('Stan', header)
    
    def test_export_csv_without_column_names(self):
        """Test CSV export without column name mapping (backward compatibility)."""
        results = [
            {
                'dbfValue': '12345',
                'stawka': '150.00',
                'czesci': 'ABC',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Numer',
                'columnIndex': 0,
                'rowIndex': 1,
                'matchedValue': '12345',
                'notes': 'Test'
            }
        ]
        
        csv_output = export_quadra_results_to_csv(results)
        lines = csv_output.strip().split('\n')
        
        # Check default header
        header = lines[0]
        self.assertIn('DBF_Value', header)
        self.assertIn('Stawka', header)
        self.assertIn('Status', header)
    
    def test_export_json_with_dict_column_names(self):
        """Test JSON export with dictionary column name mapping."""
        results = [
            {
                'dbfValue': '12345',
                'stawka': '150.00',
                'czesci': 'ABC',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Numer',
                'columnIndex': 0,
                'rowIndex': 1,
                'matchedValue': '12345',
                'notes': 'Test'
            }
        ]
        
        column_names = {
            'dbfValue': 'wartoscDBF',
            'stawka': 'cena',
            'status': 'stan'
        }
        
        json_output = export_quadra_results_to_json(results, column_names)
        
        # Check that custom keys are used
        self.assertEqual(len(json_output), 1)
        result = json_output[0]
        self.assertIn('wartoscDBF', result)
        self.assertIn('cena', result)
        self.assertIn('stan', result)
        self.assertEqual(result['wartoscDBF'], '12345')
        self.assertEqual(result['cena'], '150.00')
    
    def test_export_json_with_list_column_names(self):
        """Test JSON export with list column name mapping."""
        results = [
            {
                'dbfValue': '12345',
                'stawka': '150.00',
                'czesci': 'ABC',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Numer',
                'columnIndex': 0,
                'rowIndex': 1,
                'matchedValue': '12345',
                'notes': 'Test'
            }
        ]
        
        column_names = ['wartoscDBF', 'cena', 'stan', 'nazwaArkusza', 'nazwaKolumny',
                       'indeksKolumny', 'indeksWiersza', 'znalezionaWartosc', 
                       'czesci', 'notatki']
        
        json_output = export_quadra_results_to_json(results, column_names)
        
        # Check that custom keys are used
        self.assertEqual(len(json_output), 1)
        result = json_output[0]
        self.assertIn('wartoscDBF', result)
        self.assertIn('cena', result)
        self.assertIn('stan', result)
        self.assertEqual(result['wartoscDBF'], '12345')
        self.assertEqual(result['cena'], '150.00')
    
    def test_export_json_without_column_names(self):
        """Test JSON export without column name mapping (backward compatibility)."""
        results = [
            {
                'dbfValue': '12345',
                'stawka': '150.00',
                'czesci': 'ABC',
                'found': True,
                'sheetName': 'Sheet1',
                'columnName': 'Numer',
                'columnIndex': 0,
                'rowIndex': 1,
                'matchedValue': '12345',
                'notes': 'Test'
            }
        ]
        
        json_output = export_quadra_results_to_json(results)
        
        # Check default keys are used
        self.assertEqual(len(json_output), 1)
        result = json_output[0]
        self.assertIn('dbfValue', result)
        self.assertIn('stawka', result)
        self.assertIn('status', result)
        self.assertEqual(result['dbfValue'], '12345')
        self.assertEqual(result['stawka'], '150.00')


if __name__ == '__main__':
    unittest.main()
