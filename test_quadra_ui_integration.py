"""
test_quadra_ui_integration.py

Integration tests for Quadra tab column name mapping feature.
Tests the complete flow from settings file to UI table headers.

This test validates that the Quadra tab correctly applies column name mapping
when rendering table headers, as described in the feature specification.
"""

import unittest
import json
import os
import tempfile
from quadra_service import get_quadra_table_headers


class TestQuadraUIHeaderMappingIntegration(unittest.TestCase):
    """Integration tests for Quadra UI column header mapping."""
    
    def test_default_headers_without_mapping(self):
        """Test that default Polish headers are used when no mapping is configured."""
        # Simulate: No settings file or no quadra_column_names key
        app_settings = {}
        column_names = app_settings.get('quadra_column_names', None)
        
        # This is what happens in gui.py main() function
        headers = get_quadra_table_headers(column_names)
        
        # Expected default Polish headers
        expected = ['Arkusz', 'Płatnik', 'Numer z DBF', 'Stawka', 'Czesci', 
                    'Status', 'Kolumna', 'Wiersz', 'Uwagi']
        
        self.assertEqual(headers, expected)
    
    def test_headers_with_dictionary_mapping_from_settings(self):
        """Test headers with dictionary mapping loaded from settings file."""
        # Simulate: Settings file with quadra_column_names configuration
        app_settings = {
            'quadra_column_names': {
                'Arkusz': 'Sheet',
                'Numer z DBF': 'Order Number',
                'Stawka': 'Rate',
                'Czesci': 'Parts'
            }
        }
        column_names = app_settings.get('quadra_column_names', None)
        
        # This is what happens in gui.py main() function
        headers = get_quadra_table_headers(column_names)
        
        # Verify mapped columns
        self.assertEqual(headers[0], 'Sheet')  # Arkusz -> Sheet
        self.assertEqual(headers[2], 'Order Number')  # Numer z DBF -> Order Number
        self.assertEqual(headers[3], 'Rate')  # Stawka -> Rate
        self.assertEqual(headers[4], 'Parts')  # Czesci -> Parts
        
        # Verify unmapped columns stay as default
        self.assertEqual(headers[1], 'Płatnik')  # Unmapped
        self.assertEqual(headers[5], 'Status')  # Unmapped
        self.assertEqual(headers[6], 'Kolumna')  # Unmapped
        self.assertEqual(headers[7], 'Wiersz')  # Unmapped
        self.assertEqual(headers[8], 'Uwagi')  # Unmapped
    
    def test_headers_with_list_mapping_from_settings(self):
        """Test headers with list mapping loaded from settings file."""
        # Simulate: Settings file with list-based quadra_column_names
        app_settings = {
            'quadra_column_names': [
                'Sheet',
                'Payer', 
                'Order Number',
                'Rate',
                'Parts',
                'Status',
                'Column',
                'Row',
                'Notes'
            ]
        }
        column_names = app_settings.get('quadra_column_names', None)
        
        # This is what happens in gui.py main() function
        headers = get_quadra_table_headers(column_names)
        
        # All headers should match the provided list
        expected = ['Sheet', 'Payer', 'Order Number', 'Rate', 'Parts',
                    'Status', 'Column', 'Row', 'Notes']
        self.assertEqual(headers, expected)
    
    def test_dbf_field_name_to_display_name_mapping(self):
        """
        Test that DBF field names (NUMER, STAWKA, CZESCI) are properly
        mapped to user-friendly display names.
        
        This test addresses the issue where the Quadra UI was showing
        original DBF field names instead of translated display names.
        """
        # The default headers already provide friendly names for DBF fields
        default_headers = get_quadra_table_headers(None)
        
        # Verify that we're NOT showing raw DBF field names
        self.assertNotIn('NUMER', default_headers)
        self.assertNotIn('STAWKA', default_headers)
        self.assertNotIn('CZESCI', default_headers)
        
        # Verify that we ARE showing user-friendly Polish names
        self.assertIn('Numer z DBF', default_headers)
        self.assertIn('Stawka', default_headers)
        self.assertIn('Czesci', default_headers)
    
    def test_case_insensitive_mapping_from_settings(self):
        """Test that mapping keys are matched case-insensitively."""
        app_settings = {
            'quadra_column_names': {
                'arkusz': 'Sheet',  # lowercase
                'NUMER Z DBF': 'Order Number',  # uppercase
                'StAwKa': 'Rate'  # mixed case
            }
        }
        column_names = app_settings.get('quadra_column_names', None)
        headers = get_quadra_table_headers(column_names)
        
        # All should match despite case differences
        self.assertEqual(headers[0], 'Sheet')
        self.assertEqual(headers[2], 'Order Number')
        self.assertEqual(headers[3], 'Rate')
    
    def test_whitespace_normalization_in_mapping(self):
        """Test that mapping handles extra whitespace correctly."""
        app_settings = {
            'quadra_column_names': {
                ' Arkusz ': 'Sheet',  # extra spaces
                'Numer z DBF  ': 'Order Number',  # trailing spaces
                '  Stawka': 'Rate'  # leading spaces
            }
        }
        column_names = app_settings.get('quadra_column_names', None)
        headers = get_quadra_table_headers(column_names)
        
        # All should match despite whitespace differences
        self.assertEqual(headers[0], 'Sheet')
        self.assertEqual(headers[2], 'Order Number')
        self.assertEqual(headers[3], 'Rate')
    
    def test_fallback_to_original_for_unmapped_columns(self):
        """Test that unmapped columns fall back to original names."""
        app_settings = {
            'quadra_column_names': {
                'Arkusz': 'Sheet'
                # Only map one column
            }
        }
        column_names = app_settings.get('quadra_column_names', None)
        headers = get_quadra_table_headers(column_names)
        
        # Mapped column
        self.assertEqual(headers[0], 'Sheet')
        
        # Unmapped columns should keep original names
        self.assertEqual(headers[1], 'Płatnik')
        self.assertEqual(headers[2], 'Numer z DBF')
        self.assertEqual(headers[3], 'Stawka')
        self.assertEqual(headers[4], 'Czesci')
        self.assertEqual(headers[5], 'Status')
        self.assertEqual(headers[6], 'Kolumna')
        self.assertEqual(headers[7], 'Wiersz')
        self.assertEqual(headers[8], 'Uwagi')
    
    def test_partial_list_mapping(self):
        """Test that partial list mapping works correctly."""
        app_settings = {
            'quadra_column_names': [
                'Sheet',
                'Payer',
                'Order Number'
                # Only first 3 columns mapped
            ]
        }
        column_names = app_settings.get('quadra_column_names', None)
        headers = get_quadra_table_headers(column_names)
        
        # First 3 should be mapped
        self.assertEqual(headers[0], 'Sheet')
        self.assertEqual(headers[1], 'Payer')
        self.assertEqual(headers[2], 'Order Number')
        
        # Rest should use defaults
        self.assertEqual(headers[3], 'Stawka')
        self.assertEqual(headers[4], 'Czesci')
        self.assertEqual(headers[5], 'Status')
        self.assertEqual(headers[6], 'Kolumna')
        self.assertEqual(headers[7], 'Wiersz')
        self.assertEqual(headers[8], 'Uwagi')
    
    def test_empty_mapping_returns_defaults(self):
        """Test that empty mapping configurations return default headers."""
        # Empty dict
        headers1 = get_quadra_table_headers({})
        expected = ['Arkusz', 'Płatnik', 'Numer z DBF', 'Stawka', 'Czesci', 
                    'Status', 'Kolumna', 'Wiersz', 'Uwagi']
        self.assertEqual(headers1, expected)
        
        # Empty list
        headers2 = get_quadra_table_headers([])
        self.assertEqual(headers2, expected)


class TestQuadraUIEndToEndFlow(unittest.TestCase):
    """End-to-end integration tests simulating the complete GUI flow."""
    
    def test_complete_flow_from_settings_to_table(self):
        """
        Test the complete flow:
        1. Load settings from file
        2. Extract quadra_column_names
        3. Create Quadra tab with mapped headers
        4. Verify headers match mapping
        """
        # Step 1: Simulate settings file content
        settings_content = {
            'quadra_column_names': {
                'Arkusz': 'Sheet Name',
                'Płatnik': 'Payer Name',
                'Numer z DBF': 'DBF Order Number',
                'Stawka': 'Unit Rate',
                'Czesci': 'Parts List',
                'Status': 'Order Status',
                'Kolumna': 'Column Location',
                'Wiersz': 'Row Number',
                'Uwagi': 'Additional Notes'
            }
        }
        
        # Step 2: Extract column names (as in gui.py main())
        quadra_column_names = settings_content.get('quadra_column_names', None)
        
        # Step 3: Create Quadra tab headers (as in gui.py create_quadra_tab())
        table_headings = get_quadra_table_headers(quadra_column_names)
        
        # Step 4: Verify all mappings applied correctly
        expected_headers = [
            'Sheet Name',
            'Payer Name',
            'DBF Order Number',
            'Unit Rate',
            'Parts List',
            'Order Status',
            'Column Location',
            'Row Number',
            'Additional Notes'
        ]
        
        self.assertEqual(table_headings, expected_headers)
        
        # Verify no original names remain
        self.assertNotIn('Arkusz', table_headings)
        self.assertNotIn('Płatnik', table_headings)
        self.assertNotIn('Numer z DBF', table_headings)
        self.assertNotIn('Stawka', table_headings)
        self.assertNotIn('Czesci', table_headings)
    
    def test_backward_compatibility_without_settings(self):
        """
        Test backward compatibility: when no settings file exists,
        the application should work with default Polish headers.
        """
        # Simulate: No settings file exists
        app_settings = {}
        quadra_column_names = app_settings.get('quadra_column_names', None)
        
        # Create Quadra tab headers
        table_headings = get_quadra_table_headers(quadra_column_names)
        
        # Should get default Polish headers
        expected = ['Arkusz', 'Płatnik', 'Numer z DBF', 'Stawka', 'Czesci', 
                    'Status', 'Kolumna', 'Wiersz', 'Uwagi']
        self.assertEqual(table_headings, expected)


if __name__ == '__main__':
    unittest.main()
