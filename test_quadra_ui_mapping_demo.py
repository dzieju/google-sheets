#!/usr/bin/env python3
"""
test_quadra_ui_mapping_demo.py
Demonstration of Quadra tab column name mapping feature.
"""

import json
import os
from quadra_service import get_quadra_table_headers

def demo_default_headers():
    """Demonstrate default headers (no mapping)."""
    print("=" * 70)
    print("DEMO 1: Default Headers (No Mapping)")
    print("=" * 70)
    
    headers = get_quadra_table_headers(None)
    print("\nHeaders (Polish default):")
    for i, header in enumerate(headers, 1):
        print(f"  {i}. {header}")
    print()


def demo_dict_mapping():
    """Demonstrate dictionary-based mapping."""
    print("=" * 70)
    print("DEMO 2: Dictionary Mapping (English names)")
    print("=" * 70)
    
    # Map Polish headers to English
    column_names = {
        "Arkusz": "Sheet",
        "Płatnik": "Payer",
        "Numer z DBF": "Order Number",
        "Stawka": "Rate",
        "Czesci": "Parts",
        "Status": "Status",
        "Kolumna": "Column",
        "Wiersz": "Row",
        "Uwagi": "Notes"
    }
    
    headers = get_quadra_table_headers(column_names)
    print("\nHeaders (English):")
    for i, header in enumerate(headers, 1):
        print(f"  {i}. {header}")
    print()


def demo_partial_dict_mapping():
    """Demonstrate partial dictionary mapping."""
    print("=" * 70)
    print("DEMO 3: Partial Dictionary Mapping (Only some columns)")
    print("=" * 70)
    
    # Only map a few columns
    column_names = {
        "Arkusz": "Zakładka",
        "Numer z DBF": "Nr zlecenia",
        "Stawka": "Cena"
    }
    
    headers = get_quadra_table_headers(column_names)
    print("\nHeaders (mixed Polish with custom names):")
    for i, header in enumerate(headers, 1):
        print(f"  {i}. {header}")
    print()


def demo_list_mapping():
    """Demonstrate list-based mapping."""
    print("=" * 70)
    print("DEMO 4: List Mapping (All columns in order)")
    print("=" * 70)
    
    # Provide all 9 column names in order
    column_names = [
        "Karta",
        "Płatnik",
        "Numer",
        "Kwota",
        "Części",
        "Stan",
        "Kolumna",
        "Wiersz",
        "Notatki"
    ]
    
    headers = get_quadra_table_headers(column_names)
    print("\nHeaders (custom Polish names):")
    for i, header in enumerate(headers, 1):
        print(f"  {i}. {header}")
    print()


def demo_case_insensitive_mapping():
    """Demonstrate case-insensitive mapping."""
    print("=" * 70)
    print("DEMO 5: Case-Insensitive Mapping")
    print("=" * 70)
    
    # Use different cases in mapping keys
    column_names = {
        "arkusz": "SHEET",           # lowercase key
        "PŁATNIK": "payer",          # uppercase key
        "nUmEr Z dBf": "Order#"      # mixed case key
    }
    
    headers = get_quadra_table_headers(column_names)
    print("\nHeaders (case-insensitive matching):")
    for i, header in enumerate(headers, 1):
        print(f"  {i}. {header}")
    print()


def demo_settings_file_format():
    """Show example settings file format."""
    print("=" * 70)
    print("DEMO 6: Settings File Format")
    print("=" * 70)
    
    settings_example = {
        "quadra_column_names": {
            "Arkusz": "Sheet",
            "Płatnik": "Payer",
            "Numer z DBF": "Order Number",
            "Stawka": "Rate",
            "Czesci": "Parts",
            "Status": "Status",
            "Kolumna": "Column",
            "Wiersz": "Row",
            "Uwagi": "Notes"
        }
    }
    
    print("\nExample ~/.google_sheets_settings.json:")
    print(json.dumps(settings_example, ensure_ascii=False, indent=2))
    print("\nOr with list format:")
    
    settings_example_list = {
        "quadra_column_names": [
            "Sheet", "Payer", "Order Number", "Rate", "Parts",
            "Status", "Column", "Row", "Notes"
        ]
    }
    print(json.dumps(settings_example_list, ensure_ascii=False, indent=2))
    print()


def main():
    """Run all demos."""
    print("\n" + "=" * 70)
    print("QUADRA TAB COLUMN NAME MAPPING - DEMONSTRATION")
    print("=" * 70)
    print()
    
    demo_default_headers()
    demo_dict_mapping()
    demo_partial_dict_mapping()
    demo_list_mapping()
    demo_case_insensitive_mapping()
    demo_settings_file_format()
    
    print("=" * 70)
    print("END OF DEMONSTRATION")
    print("=" * 70)
    print("\nTo use this feature in the GUI:")
    print("1. Create/edit file: ~/.google_sheets_settings.json")
    print("2. Add 'quadra_column_names' key with your mapping")
    print("3. Restart the GUI application")
    print("4. The Quadra tab table will show your custom column names")
    print()


if __name__ == "__main__":
    main()
