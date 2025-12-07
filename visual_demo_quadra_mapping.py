#!/usr/bin/env python3
"""
Visual demonstration of Quadra tab column mapping.
Shows before and after headers with custom mapping.
"""

import json
from quadra_service import get_quadra_table_headers

def print_table_headers(headers, title):
    """Print headers in a nice table format."""
    print(f"\n{title}")
    print("=" * 80)
    print("│ " + " │ ".join(f"{h:^12}" for h in headers) + " │")
    print("=" * 80)

def main():
    print("\n" + "=" * 80)
    print("QUADRA TAB COLUMN MAPPING - VISUAL DEMONSTRATION")
    print("=" * 80)
    
    # BEFORE: Default Polish headers
    default_headers = get_quadra_table_headers(None)
    print_table_headers(default_headers, "BEFORE: Default Polish Headers (No Mapping)")
    
    # AFTER: English headers with dictionary mapping
    english_mapping = {
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
    english_headers = get_quadra_table_headers(english_mapping)
    print_table_headers(english_headers, "AFTER: English Headers (Dictionary Mapping)")
    
    # Show configuration
    print("\n\nCONFIGURATION")
    print("=" * 80)
    print("To use custom column names in the Quadra tab:")
    print("\n1. Create or edit: ~/.google_sheets_settings.json")
    print("\n2. Add the following configuration:")
    print("\n```json")
    config = {"quadra_column_names": english_mapping}
    print(json.dumps(config, ensure_ascii=False, indent=2))
    print("```")
    print("\n3. Restart the GUI application")
    print("\n4. The Quadra tab will display the custom column names!")
    
    # Example with partial mapping
    partial_mapping = {
        "Numer z DBF": "Order #",
        "Stawka": "Price",
        "Uwagi": "Comments"
    }
    partial_headers = get_quadra_table_headers(partial_mapping)
    print_table_headers(partial_headers, "\n\nEXAMPLE: Partial Mapping (Only 3 columns customized)")
    
    print("\n\nKEY FEATURES")
    print("=" * 80)
    print("✓ Case-insensitive matching: 'arkusz' = 'Arkusz' = 'ARKUSZ'")
    print("✓ Whitespace tolerant: ' Arkusz ' = 'Arkusz'")
    print("✓ Partial mapping: Only customize columns you want to change")
    print("✓ Full backward compatibility: Works without configuration")
    print("✓ Consistent with CSV/JSON exports: Same mapping logic")
    print("=" * 80)
    print()

if __name__ == "__main__":
    main()
