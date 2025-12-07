"""
test_quadra_column_mapping.py
Unit tests for Quadra tab column name mapping feature.
Tests the get_quadra_table_headers() function and column mapping logic.
"""

import unittest
from quadra_service import get_quadra_table_headers, map_column_names


class TestGetQuadraTableHeaders(unittest.TestCase):
    """Tests for get_quadra_table_headers() function."""
    
    def test_default_headers_no_mapping(self):
        """Test that default Polish headers are returned when no mapping is provided."""
        headers = get_quadra_table_headers(None)
        
        expected = ['Arkusz', 'Płatnik', 'Numer z DBF', 'Stawka', 'Czesci', 
                    'Status', 'Kolumna', 'Wiersz', 'Uwagi']
        self.assertEqual(headers, expected)
    
    def test_dict_mapping_exact_match(self):
        """Test dictionary mapping with exact matches."""
        column_names = {
            'Arkusz': 'Sheet',
            'Stawka': 'Rate',
            'Status': 'State'
        }
        
        headers = get_quadra_table_headers(column_names)
        
        self.assertEqual(headers[0], 'Sheet')      # Arkusz -> Sheet
        self.assertEqual(headers[1], 'Płatnik')    # Unmapped, stays as is
        self.assertEqual(headers[3], 'Rate')       # Stawka -> Rate
        self.assertEqual(headers[5], 'State')      # Status -> State
        self.assertEqual(headers[8], 'Uwagi')      # Unmapped, stays as is
    
    def test_dict_mapping_case_insensitive(self):
        """Test dictionary mapping with case-insensitive matching."""
        column_names = {
            'arkusz': 'Sheet',           # lowercase
            'STAWKA': 'Rate',            # uppercase
            'StAtUs': 'State'            # mixed case
        }
        
        headers = get_quadra_table_headers(column_names)
        
        self.assertEqual(headers[0], 'Sheet')      # arkusz -> Arkusz -> Sheet
        self.assertEqual(headers[3], 'Rate')       # STAWKA -> Stawka -> Rate
        self.assertEqual(headers[5], 'State')      # StAtUs -> Status -> State
    
    def test_dict_mapping_whitespace_normalization(self):
        """Test dictionary mapping with whitespace normalization."""
        column_names = {
            ' Arkusz ': 'Sheet',         # extra spaces
            'Numer z DBF': 'Number',     # exact match with spaces
            ' Stawka  ': 'Rate'          # multiple spaces
        }
        
        headers = get_quadra_table_headers(column_names)
        
        self.assertEqual(headers[0], 'Sheet')      # Arkusz (trimmed)
        self.assertEqual(headers[2], 'Number')     # Numer z DBF
        self.assertEqual(headers[3], 'Rate')       # Stawka (trimmed)
    
    def test_list_mapping_all_columns(self):
        """Test list mapping with all 9 column names."""
        column_names = [
            'Sheet', 'Payer', 'Number', 'Rate', 'Parts',
            'State', 'Column', 'Row', 'Notes'
        ]
        
        headers = get_quadra_table_headers(column_names)
        
        self.assertEqual(headers, column_names)
    
    def test_list_mapping_partial(self):
        """Test list mapping with fewer than 9 column names."""
        column_names = ['Sheet', 'Payer', 'Number']  # Only first 3
        
        headers = get_quadra_table_headers(column_names)
        
        # First 3 should be mapped
        self.assertEqual(headers[0], 'Sheet')
        self.assertEqual(headers[1], 'Payer')
        self.assertEqual(headers[2], 'Number')
        # Rest should use defaults
        self.assertEqual(headers[3], 'Stawka')
        self.assertEqual(headers[4], 'Czesci')
        self.assertEqual(headers[5], 'Status')
    
    def test_empty_dict_mapping(self):
        """Test that empty dictionary returns default headers."""
        headers = get_quadra_table_headers({})
        
        expected = ['Arkusz', 'Płatnik', 'Numer z DBF', 'Stawka', 'Czesci', 
                    'Status', 'Kolumna', 'Wiersz', 'Uwagi']
        self.assertEqual(headers, expected)
    
    def test_empty_list_mapping(self):
        """Test that empty list returns default headers."""
        headers = get_quadra_table_headers([])
        
        expected = ['Arkusz', 'Płatnik', 'Numer z DBF', 'Stawka', 'Czesci', 
                    'Status', 'Kolumna', 'Wiersz', 'Uwagi']
        self.assertEqual(headers, expected)


class TestMapColumnNames(unittest.TestCase):
    """Tests for map_column_names() function with normalization."""
    
    def test_dict_mapping_with_normalization(self):
        """Test dictionary mapping uses normalization for key matching."""
        original = ['First Name', 'Last Name', 'Email']
        mapping = {
            'first name': 'Given Name',   # lowercase with space
            'LAST_NAME': 'Family Name',   # uppercase with underscore
            'email': 'E-mail'             # different case
        }
        
        result = map_column_names(original, mapping)
        
        # Should match case-insensitively with normalization
        self.assertEqual(result[0], 'Given Name')
        self.assertEqual(result[1], 'Family Name')
        self.assertEqual(result[2], 'E-mail')
    
    def test_dict_mapping_preserves_unmapped(self):
        """Test that unmapped columns are preserved in original form."""
        original = ['Name', 'Age', 'City', 'Country']
        mapping = {
            'Name': 'Full Name',
            'Age': 'Years'
        }
        
        result = map_column_names(original, mapping)
        
        self.assertEqual(result[0], 'Full Name')
        self.assertEqual(result[1], 'Years')
        self.assertEqual(result[2], 'City')      # Preserved
        self.assertEqual(result[3], 'Country')   # Preserved
    
    def test_list_mapping_partial_coverage(self):
        """Test list mapping with partial coverage."""
        original = ['A', 'B', 'C', 'D', 'E']
        mapping = ['Alpha', 'Beta', 'Gamma']
        
        result = map_column_names(original, mapping)
        
        self.assertEqual(result[0], 'Alpha')
        self.assertEqual(result[1], 'Beta')
        self.assertEqual(result[2], 'Gamma')
        self.assertEqual(result[3], 'D')  # Original
        self.assertEqual(result[4], 'E')  # Original
    
    def test_invalid_mapping_type_returns_original(self):
        """Test that invalid mapping type returns original columns."""
        original = ['A', 'B', 'C']
        mapping = "invalid"  # String instead of dict or list
        
        result = map_column_names(original, mapping)
        
        self.assertEqual(result, original)
    
    def test_none_mapping_returns_original(self):
        """Test that None mapping returns original columns unchanged."""
        original = ['Column1', 'Column2', 'Column3']
        
        result = map_column_names(original, None)
        
        self.assertEqual(result, original)


class TestQuadraColumnMappingIntegration(unittest.TestCase):
    """Integration tests for Quadra column mapping feature."""
    
    def test_consistent_with_export_mapping(self):
        """Test that GUI headers use same mapping logic as CSV/JSON export."""
        # This test verifies that get_quadra_table_headers uses the same
        # map_column_names function as export_quadra_results_to_csv/json
        
        mapping = {
            'Arkusz': 'Sheet Name',
            'Numer z DBF': 'Order Number',
            'Stawka': 'Price'
        }
        
        headers = get_quadra_table_headers(mapping)
        
        # Verify specific mappings
        self.assertIn('Sheet Name', headers)
        self.assertIn('Order Number', headers)
        self.assertIn('Price', headers)
        
        # Verify unmapped headers are preserved
        self.assertIn('Płatnik', headers)
        self.assertIn('Status', headers)
    
    def test_backward_compatibility(self):
        """Test that default behavior is preserved (backward compatible)."""
        # When no mapping is provided, should return default Polish headers
        headers = get_quadra_table_headers()
        
        default_headers = ['Arkusz', 'Płatnik', 'Numer z DBF', 'Stawka', 'Czesci', 
                          'Status', 'Kolumna', 'Wiersz', 'Uwagi']
        
        self.assertEqual(headers, default_headers)


if __name__ == '__main__':
    unittest.main()
