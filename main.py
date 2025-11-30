"""
main.py
Prosty CLI do listowania i przeszukiwania arkuszy.
"""

import argparse
import json
from google_auth import build_services
from sheets_search import list_spreadsheets_owned_by_me, search_in_spreadsheets

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

    args = p.parse_args()
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
