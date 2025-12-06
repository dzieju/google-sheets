"""
demo_multi_column.py
Demonstracja funkcjonalności wyszukiwania wielu kolumn o tej samej nazwie.

Ten skrypt pokazuje jak działa nowa funkcjonalność bez potrzeby
łączenia się z prawdziwym Google Sheets API.
"""

from sheets_search import (
    find_all_column_indices_by_name,
    normalize_header_name,
)


def demo_basic_functionality():
    """Demonstracja podstawowej funkcjonalności."""
    print("=" * 70)
    print("DEMO: Wyszukiwanie wielu kolumn o tej samej nazwie")
    print("=" * 70)
    print()
    
    # Przykład 1: Pojedyncze dopasowanie (zachowanie jak wcześniej)
    print("1. Pojedyncze dopasowanie:")
    headers1 = ["Imię", "Nazwisko", "Email", "Telefon"]
    result1 = find_all_column_indices_by_name(headers1, "Email")
    print(f"   Nagłówki: {headers1}")
    print(f"   Szukana kolumna: 'Email'")
    print(f"   Znalezione indeksy: {result1}")
    print(f"   Znalezione kolumny: {[headers1[i] for i in result1]}")
    print()
    
    # Przykład 2: Wiele dopasowań
    print("2. Wiele dopasowań - ta sama nazwa:")
    headers2 = ["Zlecenie", "Stawka", "Zlecenie", "Uwagi", "zlecenie"]
    result2 = find_all_column_indices_by_name(headers2, "Zlecenie")
    print(f"   Nagłówki: {headers2}")
    print(f"   Szukana kolumna: 'Zlecenie'")
    print(f"   Znalezione indeksy: {result2}")
    print(f"   Znalezione kolumny: {[headers2[i] for i in result2]}")
    print()
    
    # Przykład 3: Case-insensitive
    print("3. Case-insensitive (ignorowanie wielkości liter):")
    headers3 = ["NUMER", "numer", "Numer", "test"]
    result3 = find_all_column_indices_by_name(headers3, "numer")
    print(f"   Nagłówki: {headers3}")
    print(f"   Szukana kolumna: 'numer'")
    print(f"   Znalezione indeksy: {result3}")
    print(f"   Znalezione kolumny: {[headers3[i] for i in result3]}")
    print()
    
    # Przykład 4: Ignorowanie białych znaków
    print("4. Ignorowanie białych znaków:")
    headers4 = ["Numer zlecenia", " Numer zlecenia ", "Numer  zlecenia", "Test"]
    result4 = find_all_column_indices_by_name(headers4, "Numer zlecenia")
    print(f"   Nagłówki: {headers4}")
    print(f"   Szukana kolumna: 'Numer zlecenia'")
    print(f"   Znalezione indeksy: {result4}")
    print(f"   Znalezione kolumny: {[headers4[i] for i in result4]}")
    print()
    
    # Przykład 5: Normalizacja podkreślników
    print("5. Normalizacja podkreślników (_ → spacja):")
    headers5 = ["numer_zlecenia", "numer zlecenia", "numer-zlecenia", "Test"]
    result5 = find_all_column_indices_by_name(headers5, "numer_zlecenia")
    print(f"   Nagłówki: {headers5}")
    print(f"   Szukana kolumna: 'numer_zlecenia'")
    print(f"   Znalezione indeksy: {result5}")
    print(f"   Znalezione kolumny: {[headers5[i] for i in result5]}")
    print()


def demo_normalization():
    """Demonstracja normalizacji nagłówków."""
    print("=" * 70)
    print("DEMO: Normalizacja nazw nagłówków")
    print("=" * 70)
    print()
    
    test_cases = [
        "Test",
        "TEST",
        " Test ",
        "Test_Column",
        "Test  Multiple  Spaces",
        "test_with_underscores",
        None,
        123,
        45.67,
    ]
    
    for test in test_cases:
        normalized = normalize_header_name(test)
        print(f"   Input: {repr(test):30} → Output: {repr(normalized)}")
    print()


def demo_real_world_scenario():
    """Demonstracja rzeczywistego scenariusza użycia."""
    print("=" * 70)
    print("DEMO: Rzeczywisty scenariusz - wiele zakładek z duplikowanymi kolumnami")
    print("=" * 70)
    print()
    
    # Symulacja danych z różnych zakładek
    sheets_data = {
        "Styczeń 2024": ["Data", "Zlecenie", "Stawka", "Zlecenie", "Status"],
        "Luty 2024": ["Data", "Zlecenie", "Kwota", "Uwagi"],
        "Marzec 2024": ["Zlecenie", "Stawka", "Zlecenie", "Zlecenie", "Komentarz"],
    }
    
    search_column = "Zlecenie"
    
    print(f"Szukamy kolumny: '{search_column}' we wszystkich zakładkach")
    print()
    
    total_columns_found = 0
    for sheet_name, headers in sheets_data.items():
        indices = find_all_column_indices_by_name(headers, search_column)
        print(f"📊 Zakładka: {sheet_name}")
        print(f"   Nagłówki: {headers}")
        print(f"   Znalezione indeksy kolumn: {indices}")
        print(f"   Liczba kolumn: {len(indices)}")
        
        if indices:
            for idx in indices:
                col_letter = chr(65 + idx)  # A=0, B=1, etc.
                print(f"   ✓ Kolumna {col_letter}: '{headers[idx]}'")
        else:
            print(f"   ✗ Nie znaleziono kolumny '{search_column}'")
        
        print()
        total_columns_found += len(indices)
    
    print(f"PODSUMOWANIE: Znaleziono {total_columns_found} kolumn '{search_column}' w {len(sheets_data)} zakładkach")
    print()


def demo_comparison():
    """Porównanie starego i nowego zachowania."""
    print("=" * 70)
    print("DEMO: Porównanie starego i nowego zachowania")
    print("=" * 70)
    print()
    
    headers = ["Zlecenie", "Stawka", "Zlecenie", "Uwagi", "ZLECENIE"]
    
    print(f"Nagłówki: {headers}")
    print()
    
    # Stare zachowanie (tylko pierwszy indeks)
    from sheets_search import find_column_index_by_name
    old_result = find_column_index_by_name(headers, "Zlecenie")
    print("STARE ZACHOWANIE (find_column_index_by_name):")
    print(f"   Zwrócony indeks: {old_result}")
    if old_result is not None:
        print(f"   Znaleziona kolumna: '{headers[old_result]}' (tylko pierwsza!)")
    print()
    
    # Nowe zachowanie (wszystkie indeksy)
    new_result = find_all_column_indices_by_name(headers, "Zlecenie")
    print("NOWE ZACHOWANIE (find_all_column_indices_by_name):")
    print(f"   Zwrócone indeksy: {new_result}")
    print(f"   Znalezione kolumny: {[headers[i] for i in new_result]}")
    print(f"   Liczba znalezionych kolumn: {len(new_result)}")
    print()
    
    print("💡 KORZYŚCI:")
    print("   - Wyszukiwanie znajduje WSZYSTKIE kolumny o podanej nazwie")
    print("   - Obsługa wielu kolumn w jednej zakładce")
    print("   - Obsługa wielu zakładek z tą samą nazwą kolumny")
    print("   - Zachowana kompatybilność wsteczna (stara funkcja nadal działa)")
    print()


if __name__ == "__main__":
    demo_basic_functionality()
    print()
    
    demo_normalization()
    print()
    
    demo_real_world_scenario()
    print()
    
    demo_comparison()
    print()
    
    print("=" * 70)
    print("Koniec demonstracji")
    print("=" * 70)
