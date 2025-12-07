#!/usr/bin/env python3
"""
compare_sheet_dbf.py

Uwaga:
- Wymaga pliku z kluczem konta serwisowego Google (JSON) z dostępem do odczytu arkusza.
- Udostępnij arkusz kontu serwisowemu (email z pliku JSON) albo użyj kredencji użytkownika.
- Instalacja zależności:
    pip install gspread oauth2client pandas dbfread python-dotenv

Przykład uruchomienia:
    python3 scripts/compare_sheet_dbf.py \
      --creds service-account.json \
      --sheet-id 1AbC...xyz \
      --worksheet "Marcin" \
      --column "Pazdziernik" \
      --dbf path/to/file.dbf \
      --key "Zlecenie" \
      --out-dir reports

Skrypt pozwala też zamiast kolumny podać zakres A1 (np. "A1:E100") przez --range.
"""

import argparse
import os
import sys
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dbfread import DBF


def read_google_sheet(creds_path, sheet_id, worksheet=None, range_a1=None):
    scope = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheet_id)
    if worksheet:
        ws = sh.worksheet(worksheet)
    else:
        ws = sh.sheet1

    if range_a1:
        values = ws.get(range_a1)
        if not values:
            return pd.DataFrame()
        # zakładamy pierwszy wiersz to nagłówki jeśli jest >1 wiersz
        if len(values) >= 2:
            header = values[0]
            rows = values[1:]
            df = pd.DataFrame(rows, columns=header)
        else:
            # tylko nagłówki lub tylko jeden wiersz - zwracamy jako jednowierszowy DF
            df = pd.DataFrame(values)
    else:
        # pobierz wszystkie rekordy (z nagłówkami)
        df = pd.DataFrame(ws.get_all_records())
    return df


def read_dbf(dbf_path, encoding=None):
    # dbfread zwraca iterator słowników
    table = DBF(dbf_path, encoding=encoding, load=True)
    df = pd.DataFrame(list(table))
    return df


def normalize_cols(df):
    # usuwamy nadmiarowe spacje w nazwach kolumn i robimy str
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def compare_data(sheet_df, dbf_df, key):
    sheet = normalize_cols(sheet_df)
    dbf = normalize_cols(dbf_df)

    if key not in sheet.columns:
        raise KeyError(f'Klucz "{key}" nie znaleziony w arkuszu. Dostępne kolumny: {list(sheet.columns)}')
    if key not in dbf.columns:
        raise KeyError(f'Klucz "{key}" nie znaleziony w pliku DBF. Dostępne kolumny: {list(dbf.columns)}')

    # ujednolicenie typów na str dla klucza
    sheet[key] = sheet[key].astype(str).str.strip()
    dbf[key] = dbf[key].astype(str).str.strip()

    set_sheet = set(sheet[key].dropna().unique())
    set_dbf = set(dbf[key].dropna().unique())

    only_in_sheet = sorted(list(set_sheet - set_dbf))
    only_in_dbf = sorted(list(set_dbf - set_sheet))
    in_both = sorted(list(set_sheet & set_dbf))

    # szczegółowe porównanie dla wspólnych kluczy
    merged = pd.merge(
        sheet, dbf,
        on=key,
        how='inner',
        suffixes=('_sheet', '_dbf'),
        indicator=False
    )

    # porównaj pozostałe kolumny: znajdź te w których są różnice
    diffs = []
    for _, row in merged.iterrows():
        key_val = row[key]
        row_diff = {'key': key_val}
        different = False
        # zidentyfikuj pary kolumn oryginalnych (bez sufiksów) - szukamy kolumn które występują w obu ale z różnymi nazwami
        # weźmy listę unikalnych podstawowych nazw kolumn po usunięciu sufiksów
        cols_sheet = [c for c in merged.columns if c.endswith('_sheet')]
        for col_sheet in cols_sheet:
            base = col_sheet[:-6]  # usuń _sheet
            col_dbf = base + '_dbf'
            val_sheet = row.get(col_sheet)
            val_dbf = row.get(col_dbf) if col_dbf in merged.columns else None
            # Porównanie: traktuj NaN jako puste
            s = '' if pd.isna(val_sheet) else str(val_sheet).strip()
            d = '' if pd.isna(val_dbf) else str(val_dbf).strip()
            if s != d:
                different = True
                row_diff[base] = {'sheet': s, 'dbf': d}
        if different:
            diffs.append(row_diff)

    return {
        'only_in_sheet': only_in_sheet,
        'only_in_dbf': only_in_dbf,
        'value_differences': diffs,
    }


def save_reports(result, out_dir, keyname='key'):
    os.makedirs(out_dir, exist_ok=True)
    # only_in_sheet
    f1 = os.path.join(out_dir, 'only_in_sheet.csv')
    pd.DataFrame(result['only_in_sheet'], columns=[keyname]).to_csv(f1, index=False, encoding='utf-8-sig')
    f2 = os.path.join(out_dir, 'only_in_dbf.csv')
    pd.DataFrame(result['only_in_dbf'], columns=[keyname]).to_csv(f2, index=False, encoding='utf-8-sig')
    # differences: zapisz jako JSON-like table
    diffs = []
    for d in result['value_differences']:
        row = {keyname: d['key']}
        for col, vv in d.items():
            if col == 'key':
                continue
            row[f'{col}__sheet'] = vv.get('sheet', '')
            row[f'{col}__dbf'] = vv.get('dbf', '')
        diffs.append(row)
    df_diffs = pd.DataFrame(diffs)
    f3 = os.path.join(out_dir, 'differences.csv')
    df_diffs.to_csv(f3, index=False, encoding='utf-8-sig')
    return [f1, f2, f3]


def main():
    parser = argparse.ArgumentParser(description='Porównaj zlecenia z arkusza Google i pliku DBF')
    parser.add_argument('--creds', required=True, help='ścieżka do pliku JSON z Service Account')
    parser.add_argument('--sheet-id', required=True, help='ID arkusza Google (część z URL)')
    parser.add_argument('--worksheet', required=False, help='nazwa arkusza/zakładki (np. "Marcin")')
    parser.add_argument('--range', dest='range_a1', required=False, help='zakres A1 do odczytu np. "A1:E100"')
    parser.add_argument('--column', required=False, help='jeśli chcesz odczytać tylko jedną kolumnę nazwą np. "Pazdziernik"')
    parser.add_argument('--dbf', dest='dbf_path', required=True, help='ścieżka do pliku .dbf')
    parser.add_argument('--dbf-encoding', dest='dbf_encoding', required=False, default=None, help='kodowanie DBF (opcjonalnie, np. cp1250)')
    parser.add_argument('--key', required=True, help='nazwa kolumny-klucza do porównań, np. "Zlecenie"')
    parser.add_argument('--out-dir', dest='out_dir', required=False, default=None, help='folder do zapisu raportów CSV (opcjonalnie)')
    args = parser.parse_args()

    try:
        sheet_df = read_google_sheet(args.creds, args.sheet_id, worksheet=args.worksheet, range_a1=args.range_a1)
    except Exception as e:
        print(f'Błąd przy czytaniu arkusza Google: {e}', file=sys.stderr)
        sys.exit(2)

    if args.column:
        col = str(args.column).strip()
        if col not in sheet_df.columns:
            print(f'Kolumna "{col}" nie znaleziona w odczytanym arkuszu. Dostępne kolumny: {list(sheet_df.columns)}', file=sys.stderr)
            sys.exit(3)
        # zachowujemy też klucz, zakładamy że user podał --key
        sheet_df = sheet_df[[args.key, col]] if args.key in sheet_df.columns else sheet_df[[col]]
        # jeśli --key nie było w dataframe to użytkownik musi zapewnić klucz w innym miejscu
    try:
        dbf_df = read_dbf(args.dbf_path, encoding=args.dbf_encoding)
    except Exception as e:
        print(f'Błąd przy czytaniu pliku DBF: {e}', file=sys.stderr)
        sys.exit(4)

    try:
        result = compare_data(sheet_df, dbf_df, args.key)
    except KeyError as e:
        print(str(e), file=sys.stderr)
        sys.exit(5)

    # wypisz krótkie podsumowanie
    print('Porównanie zakończone.')
    print(f'Zleceń tylko w arkuszu: {len(result["only_in_sheet"])}')
    print(f'Zleceń tylko w DBF: {len(result["only_in_dbf"])}')
    print(f'Zleceń z różnicami wartości: {len(result["value_differences"])}')

    if args.out_dir:
        files = save_reports(result, args.out_dir, keyname=args.key)
        print('Zapisane raporty:')
        for f in files:
            print(' -', f)

    # także wypisz kilka przykładów różnic na konsolę (do 20)
    if result['only_in_sheet']:
        print('\nPrzykłady zleceń tylko w arkuszu (max 20):')
        for v in result['only_in_sheet'][:20]:
            print('  ', v)
    if result['only_in_dbf']:
        print('\nPrzykłady zleceń tylko w DBF (max 20):')
        for v in result['only_in_dbf'][:20]:
            print('  ', v)
    if result['value_differences']:
        print('\nPrzykłady różnic wartości (max 10):')
        for d in result['value_differences'][:10]:
            print(' -', d)

if __name__ == '__main__':
    main()
