"""
test_column_preview.py
Unit tests for column preview functionality.
"""

import unittest
from unittest.mock import MagicMock, Mock
from sheets_search import get_sheet_headers_with_indices, get_sheet_data


class TestColumnPreview(unittest.TestCase):
    """Test column preview helper functions."""
    
    def test_get_sheet_headers_with_indices_basic(self):
        """Test: get_sheet_headers_with_indices returns proper format."""
        # Mock sheets_service
        sheets_service = MagicMock()
        mock_response = {
            'values': [['Name', 'Age', 'City']]
        }
        sheets_service.spreadsheets().values().get().execute.return_value = mock_response
        
        result = get_sheet_headers_with_indices(sheets_service, 'test_id', 'Sheet1')
        
        expected = [
            {'name': 'Name', 'index': 1},
            {'name': 'Age', 'index': 2},
            {'name': 'City', 'index': 3}
        ]
        self.assertEqual(result, expected)
    
    def test_get_sheet_headers_with_indices_empty_headers(self):
        """Test: empty headers are excluded."""
        sheets_service = MagicMock()
        mock_response = {
            'values': [['Name', '', 'City', '']]
        }
        sheets_service.spreadsheets().values().get().execute.return_value = mock_response
        
        result = get_sheet_headers_with_indices(sheets_service, 'test_id', 'Sheet1')
        
        expected = [
            {'name': 'Name', 'index': 1},
            {'name': 'City', 'index': 3}
        ]
        self.assertEqual(result, expected)
    
    def test_get_sheet_headers_with_indices_no_data(self):
        """Test: no data returns empty list."""
        sheets_service = MagicMock()
        mock_response = {'values': []}
        sheets_service.spreadsheets().values().get().execute.return_value = mock_response
        
        result = get_sheet_headers_with_indices(sheets_service, 'test_id', 'Sheet1')
        
        self.assertEqual(result, [])
    
    def test_get_sheet_data_basic(self):
        """Test: get_sheet_data returns all rows."""
        sheets_service = MagicMock()
        mock_response = {
            'values': [
                ['Name', 'Age'],
                ['Alice', '30'],
                ['Bob', '25']
            ]
        }
        sheets_service.spreadsheets().values().get().execute.return_value = mock_response
        
        result = get_sheet_data(sheets_service, 'test_id', 'Sheet1')
        
        expected = [
            ['Name', 'Age'],
            ['Alice', '30'],
            ['Bob', '25']
        ]
        self.assertEqual(result, expected)
    
    def test_get_sheet_data_empty(self):
        """Test: empty sheet returns empty list."""
        sheets_service = MagicMock()
        mock_response = {'values': []}
        sheets_service.spreadsheets().values().get().execute.return_value = mock_response
        
        result = get_sheet_data(sheets_service, 'test_id', 'Sheet1')
        
        self.assertEqual(result, [])
    
    def test_get_sheet_data_exception(self):
        """Test: exception returns empty list."""
        sheets_service = MagicMock()
        sheets_service.spreadsheets().values().get().execute.side_effect = Exception("API Error")
        
        result = get_sheet_data(sheets_service, 'test_id', 'Sheet1')
        
        self.assertEqual(result, [])


if __name__ == '__main__':
    unittest.main()
