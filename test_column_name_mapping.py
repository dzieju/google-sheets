"""
test_column_name_mapping.py
Unit tests for column name mapping feature in main.py
"""

import unittest
import json
from main import parse_column_names_arg, map_result_keys


class TestParseColumnNamesArg(unittest.TestCase):
    """Tests for parsing --column-names argument."""
    
    def test_parse_json_object(self):
        """Test parsing JSON object string."""
        input_str = '{"spreadsheetId": "ID arkusza", "sheetName": "Nazwa"}'
        result = parse_column_names_arg(input_str)
        
        self.assertIsInstance(result, dict)
        self.assertEqual(result['spreadsheetId'], 'ID arkusza')
        self.assertEqual(result['sheetName'], 'Nazwa')
    
    def test_parse_json_array(self):
        """Test parsing JSON array string."""
        input_str = '["ID arkusza", "Nazwa", "Komórka"]'
        result = parse_column_names_arg(input_str)
        
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 3)
        self.assertEqual(result[0], 'ID arkusza')
        self.assertEqual(result[1], 'Nazwa')
        self.assertEqual(result[2], 'Komórka')
    
    def test_parse_comma_separated(self):
        """Test parsing comma-separated string."""
        input_str = 'ID arkusza,Nazwa,Komórka'
        result = parse_column_names_arg(input_str)
        
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 3)
        self.assertEqual(result[0], 'ID arkusza')
        self.assertEqual(result[1], 'Nazwa')
        self.assertEqual(result[2], 'Komórka')
    
    def test_parse_comma_separated_with_spaces(self):
        """Test parsing comma-separated string with spaces."""
        input_str = ' ID arkusza , Nazwa , Komórka '
        result = parse_column_names_arg(input_str)
        
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 3)
        # Should be trimmed
        self.assertEqual(result[0], 'ID arkusza')
        self.assertEqual(result[1], 'Nazwa')
        self.assertEqual(result[2], 'Komórka')
    
    def test_parse_none(self):
        """Test parsing None."""
        result = parse_column_names_arg(None)
        self.assertIsNone(result)
    
    def test_parse_empty_string(self):
        """Test parsing empty string."""
        result = parse_column_names_arg('')
        self.assertIsNone(result)
    
    def test_parse_invalid_json(self):
        """Test parsing invalid JSON without comma - should return None."""
        input_str = 'not valid json'
        result = parse_column_names_arg(input_str)
        self.assertIsNone(result)


class TestMapResultKeys(unittest.TestCase):
    """Tests for mapping result dictionary keys."""
    
    def test_map_with_dict(self):
        """Test mapping with dictionary."""
        result = {
            'spreadsheetId': '123',
            'spreadsheetName': 'Test Sheet',
            'sheetName': 'Sheet1',
            'cell': 'A1',
            'searchedValue': 'test',
            'stawka': '100'
        }
        
        column_names = {
            'spreadsheetId': 'ID',
            'spreadsheetName': 'Nazwa arkusza',
            'sheetName': 'Zakładka'
        }
        
        mapped = map_result_keys(result, column_names)
        
        self.assertIn('ID', mapped)
        self.assertIn('Nazwa arkusza', mapped)
        self.assertIn('Zakładka', mapped)
        self.assertEqual(mapped['ID'], '123')
        self.assertEqual(mapped['Nazwa arkusza'], 'Test Sheet')
        self.assertEqual(mapped['Zakładka'], 'Sheet1')
        # Unmapped keys should remain
        self.assertIn('cell', mapped)
        self.assertIn('searchedValue', mapped)
        self.assertIn('stawka', mapped)
    
    def test_map_with_list(self):
        """Test mapping with list."""
        result = {
            'spreadsheetId': '123',
            'spreadsheetName': 'Test Sheet',
            'sheetName': 'Sheet1',
            'cell': 'A1',
            'searchedValue': 'test',
            'stawka': '100'
        }
        
        column_names = ['ID', 'Nazwa', 'Zakładka', 'Komórka', 'Wartość', 'Stawka']
        
        mapped = map_result_keys(result, column_names)
        
        self.assertIn('ID', mapped)
        self.assertIn('Nazwa', mapped)
        self.assertIn('Zakładka', mapped)
        self.assertIn('Komórka', mapped)
        self.assertIn('Wartość', mapped)
        self.assertIn('Stawka', mapped)
        self.assertEqual(mapped['ID'], '123')
        self.assertEqual(mapped['Nazwa'], 'Test Sheet')
        self.assertEqual(mapped['Zakładka'], 'Sheet1')
    
    def test_map_with_list_fewer_elements(self):
        """Test mapping with list that has fewer elements."""
        result = {
            'spreadsheetId': '123',
            'spreadsheetName': 'Test Sheet',
            'sheetName': 'Sheet1',
            'cell': 'A1',
            'searchedValue': 'test',
            'stawka': '100'
        }
        
        # Only map first 3 columns
        column_names = ['ID', 'Nazwa', 'Zakładka']
        
        mapped = map_result_keys(result, column_names)
        
        self.assertIn('ID', mapped)
        self.assertIn('Nazwa', mapped)
        self.assertIn('Zakładka', mapped)
        # Rest should use original keys
        self.assertIn('cell', mapped)
        self.assertIn('searchedValue', mapped)
        self.assertIn('stawka', mapped)
    
    def test_map_with_none(self):
        """Test mapping with None - should return original."""
        result = {
            'spreadsheetId': '123',
            'spreadsheetName': 'Test Sheet'
        }
        
        mapped = map_result_keys(result, None)
        
        self.assertEqual(mapped, result)
        self.assertIn('spreadsheetId', mapped)
        self.assertIn('spreadsheetName', mapped)
    
    def test_map_preserves_extra_keys(self):
        """Test that mapping preserves keys not in default list."""
        result = {
            'spreadsheetId': '123',
            'customField': 'custom value',
            'anotherField': 'another value'
        }
        
        column_names = ['ID']
        
        mapped = map_result_keys(result, column_names)
        
        # Mapped key should be present
        self.assertIn('ID', mapped)
        # Extra keys should be preserved
        self.assertIn('customField', mapped)
        self.assertIn('anotherField', mapped)
        self.assertEqual(mapped['customField'], 'custom value')


class TestColumnNameMappingIntegration(unittest.TestCase):
    """Integration tests for column name mapping feature."""
    
    def test_json_object_mapping_flow(self):
        """Test complete flow with JSON object mapping."""
        # Parse argument
        arg = '{"spreadsheetId": "ID", "cell": "Komórka"}'
        column_names = parse_column_names_arg(arg)
        
        # Apply to result
        result = {
            'spreadsheetId': '123',
            'spreadsheetName': 'Test',
            'sheetName': 'Sheet1',
            'cell': 'A1',
            'searchedValue': 'value',
            'stawka': '100'
        }
        
        mapped = map_result_keys(result, column_names)
        
        # Verify mapping
        self.assertIn('ID', mapped)
        self.assertIn('Komórka', mapped)
        self.assertEqual(mapped['ID'], '123')
        self.assertEqual(mapped['Komórka'], 'A1')
    
    def test_json_array_mapping_flow(self):
        """Test complete flow with JSON array mapping."""
        # Parse argument
        arg = '["ID", "Nazwa", "Zakładka", "Komórka", "Wartość", "Stawka"]'
        column_names = parse_column_names_arg(arg)
        
        # Apply to result
        result = {
            'spreadsheetId': '123',
            'spreadsheetName': 'Test',
            'sheetName': 'Sheet1',
            'cell': 'A1',
            'searchedValue': 'value',
            'stawka': '100'
        }
        
        mapped = map_result_keys(result, column_names)
        
        # Verify mapping
        self.assertIn('ID', mapped)
        self.assertIn('Nazwa', mapped)
        self.assertIn('Zakładka', mapped)
        self.assertIn('Komórka', mapped)
        self.assertIn('Wartość', mapped)
        self.assertIn('Stawka', mapped)
    
    def test_comma_separated_mapping_flow(self):
        """Test complete flow with comma-separated mapping."""
        # Parse argument
        arg = 'ID,Nazwa,Zakładka,Komórka,Wartość,Stawka'
        column_names = parse_column_names_arg(arg)
        
        # Apply to result
        result = {
            'spreadsheetId': '123',
            'spreadsheetName': 'Test',
            'sheetName': 'Sheet1',
            'cell': 'A1',
            'searchedValue': 'value',
            'stawka': '100'
        }
        
        mapped = map_result_keys(result, column_names)
        
        # Verify mapping
        self.assertIn('ID', mapped)
        self.assertIn('Nazwa', mapped)
        self.assertIn('Zakładka', mapped)


if __name__ == '__main__':
    unittest.main()
