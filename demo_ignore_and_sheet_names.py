#!/usr/bin/env python3
"""
demo_ignore_and_sheet_names.py

Demonstration of the Ignoruj (ignore) field functionality and sheet name display.
Shows how ignore patterns work with wildcards and how results include sheet names.

This script demonstrates the functionality without requiring Google API credentials.
"""

from sheets_search import (
    parse_ignore_patterns,
    matches_ignore_pattern,
    find_all_column_indices_by_name,
    normalize_header_name,
)


def demo_parse_ignore_patterns():
    """Demonstrate parsing of ignore patterns from user input."""
    print("=" * 70)
    print("DEMO 1: Parsing Ignore Patterns")
    print("=" * 70)
    
    # Example 1: Comma-separated
    input1 = "temp, debug, old"
    patterns1 = parse_ignore_patterns(input1)
    print(f"\nInput: '{input1}'")
    print(f"Parsed patterns: {patterns1}")
    
    # Example 2: Mixed separators with wildcards
    input2 = "temp*, *old, *debug*\ntest_col"
    patterns2 = parse_ignore_patterns(input2)
    print(f"\nInput (multiline): '{input2.replace(chr(10), '\\n')}'")
    print(f"Parsed patterns: {patterns2}")
    
    # Example 3: Empty input
    input3 = ""
    patterns3 = parse_ignore_patterns(input3)
    print(f"\nInput: '{input3}' (empty)")
    print(f"Parsed patterns: {patterns3} (empty list - no ignoring)")


def demo_wildcard_matching():
    """Demonstrate wildcard pattern matching."""
    print("\n" + "=" * 70)
    print("DEMO 2: Wildcard Pattern Matching")
    print("=" * 70)
    
    test_headers = [
        "temporary",
        "temp_column",
        "old_temp",
        "debug_mode",
        "production",
        "test_old",
    ]
    
    patterns = ["temp*", "*old", "*debug*"]
    
    print(f"\nIgnore patterns: {patterns}")
    print(f"\nTesting against headers:")
    
    for header in test_headers:
        matches = matches_ignore_pattern(header, patterns)
        status = "IGNORED ✗" if matches else "INCLUDED ✓"
        print(f"  {header:20} -> {status}")


def demo_column_filtering():
    """Demonstrate how ignore patterns filter columns during search."""
    print("\n" + "=" * 70)
    print("DEMO 3: Column Filtering with Ignore Patterns")
    print("=" * 70)
    
    # Simulated spreadsheet headers
    headers = [
        "Numer zlecenia",      # Index 0
        "Stawka",              # Index 1
        "Numer zlecenia_old",  # Index 2
        "Numer zlecenia",      # Index 3
        "Transport",           # Index 4
        "Numer zlecenia_temp", # Index 5
    ]
    
    print(f"\nSpreadsheet headers (with duplicates):")
    for i, h in enumerate(headers):
        print(f"  Column {i}: '{h}'")
    
    # Scenario 1: No ignore patterns
    print("\n--- Scenario 1: Search 'Numer zlecenia' with NO ignore patterns ---")
    column_name = "Numer zlecenia"
    indices = find_all_column_indices_by_name(headers, column_name, ignore_patterns=None)
    print(f"Search column: '{column_name}'")
    print(f"Matched column indices: {indices}")
    print(f"Matched columns: {[headers[i] for i in indices]}")
    
    # Scenario 2: Ignore old versions
    print("\n--- Scenario 2: Search 'Numer zlecenia' IGNORING '*old' ---")
    ignore_patterns = parse_ignore_patterns("*old")
    indices = find_all_column_indices_by_name(headers, column_name, ignore_patterns=ignore_patterns)
    print(f"Search column: '{column_name}'")
    print(f"Ignore patterns: {ignore_patterns}")
    print(f"Matched column indices: {indices}")
    print(f"Matched columns: {[headers[i] for i in indices]}")
    print(f"Ignored: Column 2 ('Numer zlecenia_old')")
    
    # Scenario 3: Ignore temp and old
    print("\n--- Scenario 3: Search 'Numer zlecenia' IGNORING '*temp, *old' ---")
    ignore_patterns = parse_ignore_patterns("*temp, *old")
    indices = find_all_column_indices_by_name(headers, column_name, ignore_patterns=ignore_patterns)
    print(f"Search column: '{column_name}'")
    print(f"Ignore patterns: {ignore_patterns}")
    print(f"Matched column indices: {indices}")
    print(f"Matched columns: {[headers[i] for i in indices]}")
    print(f"Ignored: Columns 2 ('Numer zlecenia_old') and 5 ('Numer zlecenia_temp')")


def demo_result_format():
    """Demonstrate how results include sheet names."""
    print("\n" + "=" * 70)
    print("DEMO 4: Result Format with Sheet Names")
    print("=" * 70)
    
    # Simulated search results
    results = [
        {
            "spreadsheetName": "Zlecenia 2025",
            "sheetName": "2025",
            "cell": "A5",
            "searchedValue": "38960",
            "stawka": "280.00"
        },
        {
            "spreadsheetName": "Zlecenia 2025",
            "sheetName": "Archiwum",
            "cell": "A12",
            "searchedValue": "38961",
            "stawka": "310.00"
        },
        {
            "spreadsheetName": "Zlecenia 2024",
            "sheetName": "2024",
            "cell": "A8",
            "searchedValue": "38960",
            "stawka": "280.00"
        },
    ]
    
    print("\nExample search results (showing sheet names):")
    print("\n" + "-" * 70)
    print(f"{'Arkusz kalkulacyjny':<25} {'Nazwa arkusza':<15} {'Zlecenie':<10} {'Stawka':<10}")
    print("-" * 70)
    
    for result in results:
        print(f"{result['spreadsheetName']:<25} {result['sheetName']:<15} {result['searchedValue']:<10} {result['stawka']:<10}")
    
    print("\n" + "=" * 70)
    print("Note: Each result clearly shows the sheet name (e.g., '2025', 'Archiwum')")
    print("This helps identify which tab/sheet the result came from.")
    print("=" * 70)


def demo_normalization():
    """Demonstrate header name normalization."""
    print("\n" + "=" * 70)
    print("DEMO 5: Header Name Normalization")
    print("=" * 70)
    
    test_cases = [
        ("Numer zlecenia", "numer zlecenia"),
        ("NUMER ZLECENIA", "numer zlecenia"),
        ("numer_zlecenia", "numer zlecenia"),
        ("  Numer  zlecenia  ", "numer zlecenia"),
        ("Numer__zlecenia", "numer zlecenia"),
    ]
    
    print("\nNormalization ensures flexible column matching:")
    print(f"\n{'Original Header':<25} -> {'Normalized':<20}")
    print("-" * 50)
    
    for original, expected in test_cases:
        normalized = normalize_header_name(original)
        match_symbol = "✓" if normalized == expected else "✗"
        print(f"{original!r:<25} -> {normalized!r:<20} {match_symbol}")
    
    print("\nAll variations match the same column!")


def main():
    """Run all demonstrations."""
    print("\n")
    print("╔" + "=" * 68 + "╗")
    print("║" + " " * 68 + "║")
    print("║" + "  DEMONSTRATION: Ignoruj Field and Sheet Names in Results".center(68) + "║")
    print("║" + " " * 68 + "║")
    print("╚" + "=" * 68 + "╝")
    
    demo_parse_ignore_patterns()
    demo_wildcard_matching()
    demo_column_filtering()
    demo_result_format()
    demo_normalization()
    
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("""
The 'Ignoruj' field allows users to exclude columns from search results:
  • Supports multiple values (comma, semicolon, or newline separated)
  • Supports wildcards: pattern*, *pattern, *pattern*
  • Case-insensitive and whitespace-normalized matching
  • Override behavior: ignored columns won't appear even if they match the query

Sheet names are displayed in all result views:
  • Table column 'Nazwa arkusza' shows the sheet/tab name (e.g., '2025')
  • JSON export includes 'sheetName' field for each result
  • Helps users identify the source of each search result

For more information, see README.md sections:
  - 'Ignorowanie kolumn (pole "Ignoruj")'
  - 'Wyświetlanie wyników'
""")
    print("=" * 70)


if __name__ == "__main__":
    main()
