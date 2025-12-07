"""
example_column_names.py

Example demonstrating the column name mapping feature.
Shows how to use custom column names in CSV/JSON exports.
"""

from quadra_service import export_quadra_results_to_csv, export_quadra_results_to_json
import json

# Sample Quadra results data
sample_results = [
    {
        'dbfValue': '12345',
        'stawka': '150.00',
        'czesci': 'ABC-123',
        'found': True,
        'sheetName': 'Zlecenia 2025',
        'columnName': 'Numer zlecenia',
        'columnIndex': 2,
        'rowIndex': 5,
        'matchedValue': '12345',
        'notes': 'Znaleziono w pierwszej kolumnie'
    },
    {
        'dbfValue': '67890',
        'stawka': '200.50',
        'czesci': 'XYZ-456',
        'found': False,
        'sheetName': '',
        'columnName': '',
        'columnIndex': None,
        'rowIndex': None,
        'matchedValue': '',
        'notes': 'Brak w arkuszu'
    }
]


def example_dict_mapping():
    """Example: Using dictionary mapping for Polish column names."""
    print("=" * 70)
    print("PRZYKŁAD 1: Mapowanie słownikowe (wybrane kolumny)")
    print("=" * 70)
    
    # Map only specific columns to Polish
    column_names = {
        'DBF_Value': 'Numer zlecenia',
        'Stawka': 'Cena jednostkowa',
        'Status': 'Stan',
        'SheetName': 'Nazwa arkusza',
        'ColumnName': 'Nazwa kolumny',
        'Notes': 'Uwagi'
    }
    
    csv_output = export_quadra_results_to_csv(sample_results, column_names=column_names)
    print("\nEksport CSV z niestandardowymi nazwami:")
    print(csv_output)
    
    json_output = export_quadra_results_to_json(sample_results, column_names={
        'dbfValue': 'numerZlecenia',
        'stawka': 'cenaJednostkowa',
        'status': 'stan',
        'sheetName': 'nazwaArkusza'
    })
    print("Eksport JSON z niestandardowymi nazwami (pierwsze 2 wyniki):")
    print(json.dumps(json_output[:2], ensure_ascii=False, indent=2))


def example_list_mapping():
    """Example: Using list mapping for all columns."""
    print("\n" + "=" * 70)
    print("PRZYKŁAD 2: Mapowanie listowe (wszystkie kolumny w kolejności)")
    print("=" * 70)
    
    # Map all columns using a list
    column_names = [
        'Numer zlecenia',  # DBF_Value
        'Cena',             # Stawka
        'Stan',             # Status
        'Arkusz',           # SheetName
        'Kolumna',          # ColumnName
        'Indeks kolumny',   # ColumnIndex
        'Indeks wiersza',   # RowIndex
        'Znaleziona wartość', # MatchedValue
        'Części',           # Czesci
        'Uwagi'             # Notes
    ]
    
    csv_output = export_quadra_results_to_csv(sample_results, column_names=column_names)
    print("\nEksport CSV z pełnym mapowaniem:")
    print(csv_output)


def example_partial_list_mapping():
    """Example: Using list mapping for first few columns only."""
    print("\n" + "=" * 70)
    print("PRZYKŁAD 3: Częściowe mapowanie listowe (pierwsze 5 kolumn)")
    print("=" * 70)
    
    # Map only first 5 columns, rest will use defaults
    column_names = [
        'Numer',
        'Cena',
        'Status',
        'Arkusz',
        'Kolumna',  # Remaining columns will use default names
    ]
    
    csv_output = export_quadra_results_to_csv(sample_results, column_names=column_names)
    print("\nEksport CSV z częściowym mapowaniem:")
    print(csv_output)


def example_no_mapping():
    """Example: Default behavior without column name mapping."""
    print("\n" + "=" * 70)
    print("PRZYKŁAD 4: Bez mapowania (domyślne nazwy kolumn)")
    print("=" * 70)
    
    csv_output = export_quadra_results_to_csv(sample_results)
    print("\nEksport CSV z domyślnymi nazwami:")
    print(csv_output)
    
    json_output = export_quadra_results_to_json(sample_results)
    print("Eksport JSON z domyślnymi nazwami (pierwsze 2 wyniki):")
    print(json.dumps(json_output[:2], ensure_ascii=False, indent=2))


def main():
    """Run all examples."""
    print("\n" + "=" * 70)
    print("PRZYKŁADY UŻYCIA NIESTANDARDOWYCH NAZW KOLUMN")
    print("=" * 70)
    print("\nTa funkcja pozwala na dostosowanie nazw kolumn w eksportach CSV i JSON.")
    print("Możesz używać mapowania słownikowego lub listowego.\n")
    
    example_dict_mapping()
    example_list_mapping()
    example_partial_list_mapping()
    example_no_mapping()
    
    print("\n" + "=" * 70)
    print("KONIEC PRZYKŁADÓW")
    print("=" * 70)
    print("\nWięcej informacji w README.md w sekcji 'Niestandardowe nazwy kolumn'")


if __name__ == "__main__":
    main()
