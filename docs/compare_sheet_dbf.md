# Porównanie zleceń z Google Sheets i pliku DBF

Ten dokument opisuje jak użyć skryptu `scripts/compare_sheet_dbf.py`.

## Wymagania
- Python 3.8+
- Zainstalowane pakiety:
  ```
  pip install gspread oauth2client pandas dbfread python-dotenv
  ```

## Przygotowanie dostępu do Google Sheets
1. Stwórz Service Account w Google Cloud Console.
2. Pobierz JSON z kluczem i zapisz jako `service-account.json`.
3. Udostępnij arkusz (Share) temu kontu service account e-mailem (np. my-service@...gserviceaccount.com).

## Przykład użycia
Porównaj kolumnę "Pazdziernik" w arkuszu "Marcin" z plikiem DBF, gdzie kluczem porównań jest kolumna "Zlecenie":
```
python3 scripts/compare_sheet_dbf.py \
  --creds service-account.json \
  --sheet-id 1AbC...XYZ \
  --worksheet "Marcin" \
  --column "Pazdziernik" \
  --dbf data/zlecenia.dbf \
  --key "Zlecenie" \
  --out-dir reports
```

- --sheet-id: ID arkusza Google (część URL po `/d/`).
- --worksheet: nazwa zakładki (np. "Marcin").
- --column: jeśli chcesz porównać tylko jedną kolumnę (np. "Pazdziernik").
- --range: alternatywnie możesz podać zakres A1 (np. "A1:E200").
- --dbf: ścieżka do pliku .dbf.
- --key: nazwa kolumny, po której dopasowujemy wiersze (np. numer zlecenia).
- --out-dir: katalog, w którym zapisane zostaną raporty CSV (opcjonalnie).

## Wyniki
Skrypt wydrukuje podsumowanie i (jeśli --out-dir) zapisze:
- reports/only_in_sheet.csv
- reports/only_in_dbf.csv
- reports/differences.csv

Pliki CSV zawierają odpowiednio listę kluczy tylko w arkuszu, tylko w DBF oraz szczegółowe różnice wartości.
