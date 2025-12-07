"""
main.py
Prosty CLI do listowania i przeszukiwania arkuszy.
"""

import argparse
import json
import sys
from google_auth import build_services
from sheets_search import list_spreadsheets_owned_by_me, search_in_spreadsheets


def parse_column_names_arg(column_names_str):
    """
    Parse --column-names argument which can be:
    - JSON object string: '{"spreadsheetId": "ID arkusza", "sheetName": "Nazwa"}'
    - JSON array string: '["ID arkusza", "Nazwa", "Komórka"]'
    - Comma-separated list: 'ID arkusza,Nazwa,Komórka'
    
    Returns:
        Dict or List or None
    """
    if not column_names_str:
        return None
    
    # Try parsing as JSON first
    try:
        parsed = json.loads(column_names_str)
        if isinstance(parsed, (dict, list)):
            return parsed
    except (json.JSONDecodeError, ValueError):
        pass
    
    # Try parsing as comma-separated list
    if ',' in column_names_str:
        return [name.strip() for name in column_names_str.split(',')]
    
    # Invalid format
    return None


def map_result_keys(result, column_names):
    """
    Map result dictionary keys to custom column names.
    
    Args:
        result: Result dictionary with default keys
        column_names: Dict or List for mapping
    
    Returns:
        Dictionary with mapped keys
    """
    if column_names is None:
        return result
    
    # Default keys in search results
    default_keys = ['spreadsheetId', 'spreadsheetName', 'sheetName', 'cell', 
                    'searchedValue', 'stawka']
    
    if isinstance(column_names, dict):
        # Map using dictionary
        mapped = {}
        for key, value in result.items():
            mapped[column_names.get(key, key)] = value
        return mapped
    
    if isinstance(column_names, list):
        # Map using list in order
        mapped = {}
        for i, key in enumerate(default_keys):
            if key in result:
                if i < len(column_names):
                    mapped[column_names[i]] = result[key]
                else:
                    mapped[key] = result[key]
        # Include any extra keys not in default_keys
        for key, value in result.items():
            if key not in default_keys:
                mapped[key] = value
        return mapped
    
    return result


def cmd_list():
    drive, sheets = build_services()
    files = list_spreadsheets_owned_by_me(drive)
    for f in files:
        print(f"{f['name']}  ({f['id']})")
    print(f"\nRazem: {len(files)} arkuszy")

def cmd_search(args):
    drive, sheets = build_services()
    results = search_in_spreadsheets(
        drive,
        sheets,
        pattern=(args.pattern if args.regex else args.query),
        regex=args.regex,
        case_sensitive=args.case,
        max_files=args.max_files,
    )
    count = 0
    for r in results:
        # Apply column name mapping if provided
        if args.column_names:
            r = map_result_keys(r, args.column_names)
        print(json.dumps(r, ensure_ascii=False))
        count += 1
    print(f"\nZnaleziono: {count} dopasowań")

def main():
    p = argparse.ArgumentParser(description="Google Sheets search CLI")
    sub = p.add_subparsers(dest="cmd")

    p_list = sub.add_parser("list", help="Listuj arkusze należące do Ciebie")

    p_search = sub.add_parser("search", help="Przeszukaj arkusze")
    p_search.add_argument("query", nargs="?", help="Tekst do wyszukania (substring). Jeśli używasz --regex, to zignoruj to pole.")
    p_search.add_argument("--regex", action="store_true", help="Traktuj pattern jako wyrażenie regularne.")
    p_search.add_argument("--pattern", help="Pattern regex (jeśli --regex).")
    p_search.add_argument("--case", action="store_true", help="Rozróżniaj wielkość liter.")
    p_search.add_argument("--max-files", type=int, default=None, help="Maksymalna liczba plików do przeszukania (przydatne do testów).")
    p_search.add_argument("--column-names", type=str, default=None, 
                         help='Niestandardowe nazwy kolumn. Akceptuje: JSON object ({"spreadsheetId": "ID"}), JSON array (["ID", "Nazwa"]) lub listę rozdzieloną przecinkami (ID,Nazwa,Komórka).')

    args = p.parse_args()
    
    # Parse column names if provided for search command
    if args.cmd == "search" and hasattr(args, 'column_names') and args.column_names:
        args.column_names = parse_column_names_arg(args.column_names)
        if args.column_names is None:
            p_search.error("Nieprawidłowy format --column-names. Użyj JSON object, JSON array lub listy rozdzielonej przecinkami.")
    else:
        if hasattr(args, 'column_names'):
            args.column_names = None
    
    if args.cmd == "list":
        cmd_list()
    elif args.cmd == "search":
        if args.regex and not args.pattern:
            p_search.error("--regex wymaga --pattern")
        if not args.regex and not args.query:
            p_search.error("Brak zapytania. Podaj query lub użyj --regex + --pattern.")
        cmd_search(args)
    else:
        p.print_help()

if __name__ == "__main__":
    main()
