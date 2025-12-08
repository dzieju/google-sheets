#!/usr/bin/env python
"""
demo_quadra_mapping_visual.py

Visual demonstration of the Quadra column name mapping feature.
Shows how headers are transformed from default Polish names to custom mapped names.
"""

import json
from quadra_service import get_quadra_table_headers


def print_header_comparison(title, mapping_config, mapping_description):
    """Print a visual comparison of default vs mapped headers."""
    print("\n" + "=" * 80)
    print(f"  {title}")
    print("=" * 80)
    
    if mapping_description:
        print(f"\n{mapping_description}")
    
    # Get default headers
    default_headers = get_quadra_table_headers(None)
    
    # Get mapped headers
    mapped_headers = get_quadra_table_headers(mapping_config)
    
    # Print configuration if provided
    if mapping_config:
        print("\nConfiguration:")
        print(json.dumps({'quadra_column_names': mapping_config}, indent=2, ensure_ascii=False))
    
    # Print comparison table
    print("\nHeader Transformation:")
    print("-" * 80)
    print(f"{'Position':<10} {'Default (Polish)':<25} {'→':<5} {'Mapped (Custom)':<30}")
    print("-" * 80)
    
    for i, (default, mapped) in enumerate(zip(default_headers, mapped_headers)):
        changed = "✓" if default != mapped else " "
        print(f"{i+1:<10} {default:<25} {changed:>3}→  {mapped:<30}")
    
    print("-" * 80)
    print(f"Changed: {sum(1 for d, m in zip(default_headers, mapped_headers) if d != m)}/9 columns")


def main():
    """Run visual demonstrations."""
    print("\n" + "=" * 80)
    print("  QUADRA COLUMN NAME MAPPING - VISUAL DEMONSTRATION")
    print("=" * 80)
    print("\nThis demonstrates how the Quadra tab applies column name mapping")
    print("when rendering table headers.")
    
    # Demo 1: No mapping (default behavior)
    print_header_comparison(
        "Demo 1: Default Behavior (No Mapping)",
        None,
        "When no configuration is provided, default Polish headers are used."
    )
    
    # Demo 2: Dictionary mapping (partial)
    print_header_comparison(
        "Demo 2: Dictionary Mapping (Partial Translation)",
        {
            "Arkusz": "Sheet",
            "Numer z DBF": "Order Number",
            "Stawka": "Rate",
            "Czesci": "Parts"
        },
        "Dictionary mapping translates only specified columns to English.\nUnmapped columns keep their Polish names."
    )
    
    # Demo 3: Dictionary mapping (full)
    print_header_comparison(
        "Demo 3: Dictionary Mapping (Full Translation)",
        {
            "Arkusz": "Sheet",
            "Płatnik": "Payer",
            "Numer z DBF": "Order Number",
            "Stawka": "Rate",
            "Czesci": "Parts",
            "Status": "Status",
            "Kolumna": "Column",
            "Wiersz": "Row",
            "Uwagi": "Notes"
        },
        "All columns translated to English using dictionary mapping."
    )
    
    # Demo 4: List mapping
    print_header_comparison(
        "Demo 4: List Mapping (Custom Names in Order)",
        [
            "Sheet Name",
            "Payer Name",
            "DBF Order #",
            "Unit Rate",
            "Parts List",
            "Order Status",
            "Column Loc",
            "Row #",
            "Additional Notes"
        ],
        "List mapping provides custom names for all 9 columns in order."
    )
    
    # Demo 5: Case-insensitive mapping
    print_header_comparison(
        "Demo 5: Case-Insensitive Mapping",
        {
            "arkusz": "Sheet",  # lowercase
            "NUMER Z DBF": "Order Number",  # uppercase
            "StAwKa": "Rate"  # mixed case
        },
        "Mapping keys are matched case-insensitively with the default headers."
    )
    
    # Demo 6: Practical example
    print_header_comparison(
        "Demo 6: Practical Example (Business-Friendly Names)",
        {
            "Arkusz": "Spreadsheet",
            "Płatnik": "Customer",
            "Numer z DBF": "Order ID",
            "Stawka": "Unit Price",
            "Czesci": "Part Number",
            "Status": "Match Status",
            "Kolumna": "Found in Column",
            "Wiersz": "Found at Row",
            "Uwagi": "Comments"
        },
        "Business-friendly English names for all columns."
    )
    
    # Summary
    print("\n" + "=" * 80)
    print("  CONFIGURATION")
    print("=" * 80)
    print("\nTo use custom column names in the Quadra tab:")
    print("1. Create or edit: ~/.google_sheets_settings.json")
    print("2. Add the 'quadra_column_names' key with your mapping")
    print("3. Restart the application")
    print("\nExample ~/.google_sheets_settings.json:")
    print(json.dumps({
        "quadra_column_names": {
            "Arkusz": "Sheet",
            "Numer z DBF": "Order Number",
            "Stawka": "Rate"
        }
    }, indent=2, ensure_ascii=False))
    
    print("\n" + "=" * 80)
    print("  FEATURES")
    print("=" * 80)
    print("✓ Case-insensitive mapping keys")
    print("✓ Whitespace normalization")
    print("✓ Fallback to original names for unmapped columns")
    print("✓ Supports both dictionary and list formats")
    print("✓ Backward compatible (works without config file)")
    print("\n" + "=" * 80)


if __name__ == "__main__":
    main()
