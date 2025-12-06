"""
sheets_search.py
Funkcje:
- list_spreadsheets_owned_by_me(drive_service)
- search_in_spreadsheets(drive_service, sheets_service, pattern, regex=False, case_sensitive=False, max_files=None)
- search_in_sheet(drive_service, sheets_service, spreadsheet_id, sheet_name, pattern, regex=False, case_sensitive=False, search_column_name=None)
- search_in_spreadsheet(drive_service, sheets_service, spreadsheet_id, pattern, regex=False, case_sensitive=False, search_column_name=None)
- get_sheet_headers(sheets_service, spreadsheet_id, sheet_name)
- detect_header_row(values, search_column_name=None)
- find_duplicates_in_sheet(drive_service, sheets_service, spreadsheet_id, sheet_name, search_column_name, normalize=True)
- find_duplicates_across_spreadsheets(drive_service, sheets_service, spreadsheet_ids, search_column_name, normalize=True)

Poprawka: normalizacja ciągów liczbowych, aby wyszukiwanie znajdowało liczby pomimo różnego formatowania (spacje, NBSP, separatory tysięcy, przecinek/kropka).
Dodatkowa odporność na wartości None i typy numeryczne (int/float).

Nowa funkcjonalność:
- Wykrywanie nagłówków "Numer zlecenia" i "Stawka" w arkuszu
- Zwracanie searchedValue (numer zlecenia) i stawka w wynikach
- Fallback: jeśli brak nagłówków, stawka z komórki po prawej
- Obsługa parametru search_column_name: 'ALL'/'Wszystkie' przeszukuje wszystkie kolumny,
  konkretna nazwa kolumny przeszukuje tylko tę kolumnę, brak wartości przeszukuje 'numer zlecenia'
- Wykrywanie nagłówków w wierszu 1 lub 2: jeśli szukana kolumna nie znajduje się w wierszu 1,
  algorytm sprawdza również wiersz 2 i odpowiednio dostosowuje indeks początkowy danych

Obsługa wielu kolumn o tej samej nazwie (v2):
- find_all_column_indices_by_name() - znajduje wszystkie kolumny o danej nazwie (case-insensitive, trimmed)
- search_in_sheet() i search_in_spreadsheet() przeszukują WSZYSTKIE kolumny o podanej nazwie
  w każdym arkuszu/zakładce, zachowując kolejność (najpierw kolejność arkuszy, potem indeks kolumny)
- find_duplicates_in_sheet() wykrywa duplikaty w każdej kolumnie osobno i rozróżnia kolumny
  w wynikach (dodaje informację o literze kolumny gdy jest wiele kolumn o tej samej nazwie)
- Zachowana kompatybilność wsteczna: funkcje zwracają wszystkie dopasowania zamiast tylko pierwszego
"""

import logging
import re
import threading
from collections import Counter
from typing import List, Dict, Any, Generator, Optional, Union, Tuple

# Konfiguracja loggera dla modułu
logger = logging.getLogger(__name__)


# helpers
def col_index_to_a1(n: Union[int, None]) -> str:
    """Konwertuje indeks kolumny 0-based na etykietę A1 (0 -> A).
    Zwraca '?' jeśli n jest None lub nieprawidłowy.
    """
    if n is None:
        return "?"
    try:
        n = int(n)
    except (TypeError, ValueError):
        return "?"
    if n < 0:
        return "?"
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_address(row_idx: Union[int, None], col_idx: Union[int, None]) -> str:
    """Zwraca adres komórki A1. Obsługuje None/nieprawidłowe indeksy defensywnie."""
    col_label = col_index_to_a1(col_idx)
    if row_idx is None:
        return f"{col_label}?"
    try:
        row_num = int(row_idx) + 1
    except (TypeError, ValueError):
        row_num = "?"
    return f"{col_label}{row_num}"


def normalize_number_string(value: Any) -> str:
    """
    Normalizuje wartość zawierającą liczbę:
    - obsługuje None (zwraca '')
    - obsługuje typy numeryczne (int/float) - konwertuje do str
    - usuwa spacje zwykłe i NBSP, znaki cienkiej spacji itp.
    - zamienia przecinek na kropkę (dla miejsc dziesiętnych)
    - usuwa wszystko poza cyframi, kropką i minusem
    Zwraca znormalizowany ciąg, lub '' jeśli po usunięciu nic nie zostaje.
    """
    if value is None:
        return ""
    # Obsługa typów numerycznych bezpośrednio
    if isinstance(value, (int, float)):
        # Dla float, używamy str() który da np. "38960.0" lub "123.45"
        s = str(value)
    else:
        s = str(value)
    # usuń niełamiące spacje i inne biały znaki grupujące
    s = s.replace("\u00A0", "").replace("\u202F", "")
    # usuń zwykłe spacje
    s = s.replace(" ", "")
    # zamień przecinek na kropkę (np. "1,23" -> "1.23")
    s = s.replace(",", ".")
    # zostaw tylko cyfry, kropkę i minus
    s = re.sub(r"[^\d\.\-]", "", s)
    return s


# Warianty nagłówków dla rozpoznawania kolumn
ZLECENIE_HEADERS = ['numer zlecenia', 'nr zlecenia', 'nr_zlecenia', 'zlecenie', 'nr z']
STAWKA_HEADERS = ['stawka', 'stawka zł', 'stawka_pln', 'stawka netto', 'stawka_brutto']

# Blacklista nazw kolumn, które NIE powinny być używane jako źródło stawki w trybie fallback
COLUMN_BLACKLIST = ['transport', 'uwagi', 'komentarz', 'komentarze', 'notatki', 'opis', 'uwaga']

# Specjalne wartości dla search_column_name wskazujące przeszukiwanie wszystkich kolumn
ALL_COLUMNS_VALUES = ['all', 'wszystkie']


def parse_header_rows(header_rows_input: Optional[str]) -> List[int]:
    """
    Parsuje konfigurację wierszy nagłówkowych z pola "Header rows".
    
    Obsługuje wartości oddzielone przecinkami (np. "1", "1,2", "1, 2, 3").
    Zwraca listę indeksów wierszy 0-based.
    
    Args:
        header_rows_input: String z numerami wierszy (1-based) oddzielonymi przecinkami
                          lub None/pusty string (domyślnie wiersz 1)
    
    Returns:
        Lista indeksów wierszy 0-based (np. [0] dla "1", [0, 1] dla "1,2")
        Zwraca [0] jeśli input jest None lub nieprawidłowy
    
    Examples:
        >>> parse_header_rows("1")
        [0]
        >>> parse_header_rows("1,2")
        [0, 1]
        >>> parse_header_rows("1, 2, 3")
        [0, 1, 2]
        >>> parse_header_rows(None)
        [0]
        >>> parse_header_rows("")
        [0]
    """
    if not header_rows_input or not header_rows_input.strip():
        return [0]  # Default: row 1 (0-based index)
    
    indices = []
    for part in header_rows_input.split(','):
        part = part.strip()
        if part:
            try:
                row_num = int(part)
                if row_num >= 1:  # 1-based input
                    indices.append(row_num - 1)  # Convert to 0-based
            except ValueError:
                continue
    
    # If no valid indices found, return default
    return indices if indices else [0]


def combine_header_values(values: List[List[Any]], header_row_indices: List[int]) -> List[str]:
    """
    Łączy wartości z wielu wierszy nagłówkowych dla każdej kolumny.
    
    Dla każdej kolumny, bierze wartości z określonych wierszy nagłówkowych,
    łączy je spacją i normalizuje (trim + lowercase).
    
    Args:
        values: Lista wszystkich wierszy arkusza
        header_row_indices: Lista indeksów wierszy nagłówkowych (0-based)
    
    Returns:
        Lista połączonych nagłówków dla każdej kolumny (znormalizowane)
    
    Examples:
        >>> values = [["Name", "Age"], ["First Last", "Years"]]
        >>> combine_header_values(values, [0, 1])
        ['name first last', 'age years']
    """
    if not values or not header_row_indices:
        return []
    
    # Find maximum number of columns across all header rows
    max_cols = 0
    for row_idx in header_row_indices:
        # Check bounds: row_idx must be non-negative and within values range
        if row_idx >= 0 and row_idx < len(values):
            max_cols = max(max_cols, len(values[row_idx]))
    
    if max_cols == 0:
        return []
    
    # Combine header values for each column
    combined_headers = []
    for col_idx in range(max_cols):
        col_parts = []
        for row_idx in header_row_indices:
            # Check bounds: row_idx must be non-negative and within values range
            if row_idx >= 0 and row_idx < len(values):
                row = values[row_idx]
                if col_idx < len(row) and row[col_idx] is not None:
                    cell_value = str(row[col_idx]).strip()
                    if cell_value:
                        col_parts.append(cell_value)
        
        # Join with space and normalize
        combined = ' '.join(col_parts)
        combined_headers.append(normalize_header_name(combined))
    
    return combined_headers


def normalize_header_name(name: Any) -> str:
    """
    Normalizuje nazwę nagłówka kolumny:
    - zamienia '_' na ' '
    - redukuje wielokrotne spacje do pojedynczej
    - strip + lowercase
    
    Args:
        name: Wartość nagłówka (może być None, int, float lub str)
    
    Returns:
        Znormalizowana nazwa nagłówka jako lowercase string
    """
    if name is None:
        return ""
    s = str(name)
    # Zamień podkreślenia na spacje
    s = s.replace('_', ' ')
    # Zredukuj wielokrotne spacje do jednej
    s = ' '.join(s.split())
    # Strip + lowercase
    return s.strip().lower()


def extract_numeric_tokens(text: str) -> List[str]:
    """
    Wyciąga tokeny liczbowe z tekstu (np. z URL-ów).
    
    Przydatne gdy w komórce jest URL zawierający numer zlecenia,
    np. "https://example.com/order/12345" -> ["12345"]
    
    Args:
        text: Tekst do przeanalizowania
    
    Returns:
        Lista znalezionych tokenów liczbowych (jako stringi)
    """
    if not text:
        return []
    # Znajdź wszystkie ciągi cyfr (z opcjonalnym separatorem dziesiętnym)
    return re.findall(r'\d+(?:[.,]\d+)?', str(text))


def is_search_all_columns(search_column_name: Optional[str]) -> bool:
    """
    Sprawdza czy search_column_name oznacza przeszukiwanie wszystkich kolumn.
    
    Args:
        search_column_name: Nazwa kolumny lub 'ALL'/'Wszystkie'
    
    Returns:
        True jeśli należy przeszukiwać wszystkie kolumny
    """
    if search_column_name is None:
        return False
    return normalize_header_name(search_column_name) in ALL_COLUMNS_VALUES


def parse_ignore_patterns(ignore_input: Optional[str]) -> List[str]:
    """
    Parsuje pole 'Ignoruj' na listę wzorców ignorowania.
    
    Obsługuje wiele wartości oddzielonych przecinkami, średnikami lub nowymi liniami.
    Każda wartość jest normalizowana (trim + lowercase).
    
    Args:
        ignore_input: Tekst z pola Ignoruj (może być None lub pusty)
    
    Returns:
        Lista znormalizowanych wzorców ignorowania (pusta lista jeśli brak)
    
    Examples:
        >>> parse_ignore_patterns("temp, test, debug")
        ['temp', 'test', 'debug']
        >>> parse_ignore_patterns("temp*\\ntest\\n*debug")
        ['temp*', 'test', '*debug']
        >>> parse_ignore_patterns(None)
        []
    """
    if not ignore_input:
        return []
    
    # Zamień średniki i nowe linie na przecinki
    normalized = ignore_input.replace(';', ',').replace('\n', ',')
    
    # Podziel na części i wyczyść
    patterns = []
    for part in normalized.split(','):
        part = part.strip()
        if part:
            # Normalizuj pattern (lowercase, ale zachowaj wildcards)
            patterns.append(part.lower())
    
    return patterns


def matches_ignore_pattern(header_name: str, ignore_patterns: List[str]) -> bool:
    """
    Sprawdza czy nazwa nagłówka pasuje do któregokolwiek wzorca ignorowania.
    
    Obsługuje proste wildcardy:
    - "pattern*" - dopasowanie prefiksu (startsWith)
    - "*pattern" - dopasowanie sufiksu (endsWith)
    - "*pattern*" - dopasowanie podciągu (contains)
    - "pattern" - dopasowanie podciągu (substring match, case-insensitive)
    
    Porównanie jest case-insensitive i po trim.
    
    Args:
        header_name: Nazwa nagłówka kolumny do sprawdzenia
        ignore_patterns: Lista wzorców ignorowania
    
    Returns:
        True jeśli nagłówek pasuje do któregokolwiek wzorca
    
    Examples:
        >>> matches_ignore_pattern("temporary", ["temp*"])
        True
        >>> matches_ignore_pattern("debug_mode", ["*debug*"])
        True
        >>> matches_ignore_pattern("test", ["test"])
        True
        >>> matches_ignore_pattern("test_column", ["test"])  # substring match
        True
        >>> matches_ignore_pattern("production", ["temp*", "test*"])
        False
    """
    if not ignore_patterns:
        return False
    
    # Normalizuj nazwę nagłówka (trim + lowercase)
    normalized_header = normalize_header_name(header_name)
    
    if not normalized_header:
        return False
    
    for pattern in ignore_patterns:
        pattern = pattern.strip().lower()
        if not pattern:
            continue
        
        # Obsługa wildcardów
        if pattern.startswith('*') and pattern.endswith('*'):
            # *pattern* - contains
            search_term = pattern[1:-1]
            if search_term and search_term in normalized_header:
                return True
        elif pattern.startswith('*'):
            # *pattern - endsWith
            search_term = pattern[1:]
            if search_term and normalized_header.endswith(search_term):
                return True
        elif pattern.endswith('*'):
            # pattern* - startsWith
            search_term = pattern[:-1]
            if search_term and normalized_header.startswith(search_term):
                return True
        else:
            # Dopasowanie podciągu (substring match, case-insensitive)
            if pattern in normalized_header:
                return True
    
    return False


def matches_ignore_value(cell_value: str, ignore_patterns: List[str]) -> bool:
    """
    Sprawdza czy wartość komórki pasuje do któregokolwiek wzorca ignorowania.
    
    Obsługuje proste wildcardy:
    - "pattern*" - dopasowanie prefiksu (startsWith)
    - "*pattern" - dopasowanie sufiksu (endsWith)
    - "*pattern*" - dopasowanie podciągu (contains)
    - "pattern" - dopasowanie podciągu (substring match, case-insensitive)
    
    Porównanie jest case-insensitive i po trim.
    
    Args:
        cell_value: Wartość komórki do sprawdzenia
        ignore_patterns: Lista wzorców ignorowania
    
    Returns:
        True jeśli wartość pasuje do któregokolwiek wzorca
    
    Examples:
        >>> matches_ignore_value("https://example.com", ["https"])
        True
        >>> matches_ignore_value("HTTP://EXAMPLE.COM", ["https"])
        True
        >>> matches_ignore_value("some text", ["https"])
        False
        >>> matches_ignore_value("test_value", ["test*"])
        True
    """
    if not ignore_patterns:
        return False
    
    if not cell_value:
        return False
    
    # Normalizuj wartość komórki (trim + lowercase)
    normalized_value = str(cell_value).strip().lower()
    
    if not normalized_value:
        return False
    
    for pattern in ignore_patterns:
        pattern = pattern.strip().lower()
        if not pattern:
            continue
        
        # Obsługa wildcardów
        if pattern.startswith('*') and pattern.endswith('*'):
            # *pattern* - contains
            search_term = pattern[1:-1]
            if search_term and search_term in normalized_value:
                return True
        elif pattern.startswith('*'):
            # *pattern - endsWith
            search_term = pattern[1:]
            if search_term and normalized_value.endswith(search_term):
                return True
        elif pattern.endswith('*'):
            # pattern* - startsWith
            search_term = pattern[:-1]
            if search_term and normalized_value.startswith(search_term):
                return True
        else:
            # Dopasowanie podciągu (substring match, case-insensitive)
            if pattern in normalized_value:
                return True
    
    return False


def get_sheet_headers(sheets_service, spreadsheet_id: str, sheet_name: str) -> List[str]:
    """
    Pobiera nagłówki z arkusza - najpierw z wiersza 1, a jeśli pusty lub nie wygląda jak nagłówek,
    to z wiersza 2.
    
    Args:
        sheets_service: Obiekt serwisu Google Sheets API
        spreadsheet_id: ID arkusza kalkulacyjnego
        sheet_name: Nazwa zakładki
    
    Returns:
        Lista nagłówków kolumn (puste stringi dla pustych komórek)
    """
    try:
        # Pobierz pierwsze dwa wiersze, żeby sprawdzić oba
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!1:2",
            majorDimension="ROWS"
        ).execute()
        values = resp.get("values", [])
        
        if not values:
            return []
        
        # Sprawdź wiersz 1
        row1 = values[0] if len(values) > 0 else []
        if row1 and is_likely_header_row(row1):
            return [str(cell) if cell is not None else "" for cell in row1]
        
        # Jeśli wiersz 1 nie wygląda jak nagłówek, sprawdź wiersz 2
        row2 = values[1] if len(values) > 1 else []
        if row2 and is_likely_header_row(row2):
            return [str(cell) if cell is not None else "" for cell in row2]
        
        # Fallback - zwróć wiersz 1 nawet jeśli nie wygląda jak nagłówek
        return [str(cell) if cell is not None else "" for cell in row1] if row1 else []
    except Exception as e:
        logger.error(f"Błąd pobierania nagłówków z [{spreadsheet_id}] {sheet_name}: {e}")
        return []


def detect_header_row(
    values: List[List[Any]], 
    search_column_name: Optional[str] = None,
    header_row_indices: Optional[List[int]] = None
) -> tuple:
    """
    Wykrywa wiersz nagłówków w arkuszu - sprawdza wiersz 1, a jeśli nie znajdzie 
    oczekiwanego nagłówka, sprawdza wiersz 2.
    
    Gdy header_row_indices jest podane, używa tych wierszy do budowy połączonych
    nagłówków (łącząc wartości z wielu wierszy spacją).
    
    Args:
        values: Lista wierszy z arkusza
        search_column_name: Nazwa kolumny do wyszukania (opcjonalna)
            - Jeśli podana, szuka tej konkretnej kolumny
            - Jeśli None, szuka 'numer zlecenia'
        header_row_indices: Lista indeksów wierszy nagłówkowych (0-based)
            - Jeśli podana, używa tych wierszy do budowy połączonych nagłówków
            - Jeśli None, wykrywa automatycznie (wiersz 1 lub 2)
    
    Returns:
        Tuple (header_row_index, header_row, start_data_row):
            - header_row_index: Indeks wiersza nagłówków (0 lub 1) lub lista indeksów
            - header_row: Lista wartości nagłówków (połączone jeśli wiele wierszy) lub None
            - start_data_row: Indeks pierwszego wiersza z danymi (po nagłówku)
    """
    if not values:
        return None, None, 0
    
    # Jeśli podano konkretne wiersze nagłówkowe, użyj ich
    if header_row_indices is not None and len(header_row_indices) > 0:
        # Combine headers from specified rows
        combined_headers = combine_header_values(values, header_row_indices)
        if combined_headers:
            # Start data row is after the last header row
            start_data_row = max(header_row_indices) + 1
            return header_row_indices, combined_headers, start_data_row
        # If combining failed, fall through to auto-detection
    
    # Auto-detection logic (original behavior)
    row1 = values[0] if len(values) > 0 else []
    row2 = values[1] if len(values) > 1 else []
    
    # Określ nazwę kolumny do szukania
    if search_column_name is not None and not is_search_all_columns(search_column_name):
        # Szukamy konkretnej kolumny po nazwie
        target_column = search_column_name
    else:
        # Domyślnie szukamy 'numer zlecenia'
        target_column = None  # Użyjemy find_header_indices
    
    def check_row_has_target(row: List[Any]) -> bool:
        """Sprawdza czy wiersz zawiera szukaną kolumnę."""
        if not row:
            return False
        if target_column:
            return find_column_index_by_name(row, target_column) is not None
        else:
            zlecenie_idx, _ = find_header_indices(row)
            return zlecenie_idx is not None
    
    # Sprawdź wiersz 1 najpierw
    if is_likely_header_row(row1):
        if check_row_has_target(row1):
            return 0, row1, 1  # Nagłówek w wierszu 1, dane od wiersza 2 (index 1)
        # Wiersz 1 wygląda jak nagłówek, ale nie ma szukanej kolumny
        # Sprawdź wiersz 2
        if is_likely_header_row(row2) and check_row_has_target(row2):
            return 1, row2, 2  # Nagłówek w wierszu 2, dane od wiersza 3 (index 2)
        # Wróć do wiersza 1 (nawet bez szukanej kolumny)
        return 0, row1, 1
    
    # Wiersz 1 nie wygląda jak nagłówek - sprawdź wiersz 2
    if is_likely_header_row(row2):
        if check_row_has_target(row2):
            return 1, row2, 2  # Nagłówek w wierszu 2, dane od wiersza 3 (index 2)
        return 1, row2, 2  # Wiersz 2 wygląda jak nagłówek (bez szukanej kolumny)
    
    # Żaden wiersz nie wygląda jak nagłówek
    return None, None, 0


def find_header_indices(header_row: List[Any]) -> tuple:
    """
    Wyszukuje indeksy kolumn "Numer zlecenia" i "Stawka" w wierszu nagłówkowym.
    
    Args:
        header_row: Lista wartości pierwszego wiersza (potencjalnego nagłówka)
    
    Returns:
        Tuple (zlecenie_idx, stawka_idx) - indeksy lub (None, None) jeśli nie znaleziono
    """
    if not header_row:
        return None, None
    
    zlecenie_idx = None
    stawka_idx = None
    
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        cell_lower = str(cell).lower().strip()
        
        # Sprawdź warianty nagłówka "Numer zlecenia"
        if zlecenie_idx is None:
            for variant in ZLECENIE_HEADERS:
                if variant in cell_lower:
                    zlecenie_idx = idx
                    break
        
        # Sprawdź warianty nagłówka "Stawka"
        if stawka_idx is None:
            for variant in STAWKA_HEADERS:
                if variant in cell_lower:
                    stawka_idx = idx
                    break
    
    return zlecenie_idx, stawka_idx


def find_column_index_by_name(header_row: List[Any], column_name: str) -> Optional[int]:
    """
    Znajduje indeks kolumny po znormalizowanej nazwie.
    
    Args:
        header_row: Lista wartości pierwszego wiersza (nagłówka)
        column_name: Nazwa kolumny do znalezienia
    
    Returns:
        Indeks kolumny lub None jeśli nie znaleziono
    """
    if not header_row or not column_name:
        return None
    
    norm_target = normalize_header_name(column_name)
    
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        norm_cell = normalize_header_name(cell)
        if norm_cell == norm_target:
            return idx
    
    return None


def find_all_column_indices_by_name(
    header_row: List[Any], 
    column_name: str, 
    ignore_patterns: Optional[List[str]] = None
) -> List[int]:
    """
    Znajduje wszystkie indeksy kolumn pasujących do podanej nazwy (znormalizowanej, case-insensitive).
    Opcjonalnie filtruje kolumny pasujące do wzorców ignorowania.
    
    Args:
        header_row: Lista wartości pierwszego wiersza (nagłówka)
        column_name: Nazwa kolumny do znalezienia
        ignore_patterns: Opcjonalna lista wzorców ignorowania (z parse_ignore_patterns)
    
    Returns:
        Lista indeksów wszystkich kolumn pasujących do nazwy i nie pasujących do ignore_patterns
    """
    if not header_row or not column_name:
        return []
    
    norm_target = normalize_header_name(column_name)
    matching_indices = []
    
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        norm_cell = normalize_header_name(cell)
        if norm_cell == norm_target:
            # Sprawdź czy kolumna nie jest ignorowana
            if ignore_patterns and matches_ignore_pattern(str(cell), ignore_patterns):
                continue  # Pomiń ignorowane kolumny
            matching_indices.append(idx)
    
    return matching_indices


def find_stawka_column_index(header_row: List[Any]) -> Optional[int]:
    """
    Znajduje indeks kolumny 'Stawka' (sprawdzając warianty nazwy).
    
    Args:
        header_row: Lista wartości pierwszego wiersza (nagłówka)
    
    Returns:
        Indeks kolumny Stawka lub None jeśli nie znaleziono
    """
    if not header_row:
        return None
    
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        cell_lower = normalize_header_name(cell)
        for variant in STAWKA_HEADERS:
            if variant in cell_lower:
                return idx
    
    return None


def is_likely_header_row(row: List[Any]) -> bool:
    """
    Sprawdza czy wiersz wygląda jak nagłówek (zawiera tekst, nie tylko liczby).
    """
    if not row:
        return False
    
    text_cells = 0
    for cell in row:
        if cell is None:
            continue
        cell_str = str(cell).strip()
        if cell_str and not re.match(r'^[\d\.\,\-\s]+$', cell_str):
            text_cells += 1
    
    # Wiersz jest nagłówkiem jeśli ma co najmniej 2 komórki z tekstem
    return text_cells >= 2


def get_cell_value_safe(row: List[Any], idx: int) -> Optional[str]:
    """
    Bezpiecznie pobiera wartość z wiersza pod danym indeksem.
    Zwraca None jeśli indeks poza zakresem lub wartość jest None.
    """
    if idx is None or idx < 0:
        return None
    if idx >= len(row):
        return None
    val = row[idx]
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return str(val)
    return str(val)


def is_column_blacklisted(header_row: Optional[List[Any]], col_idx: int) -> bool:
    """
    Sprawdza czy kolumna o danym indeksie ma nagłówek z blacklisty.
    Używane w trybie fallback, aby nie zwracać wartości z kolumn typu 'Transport'.
    
    Args:
        header_row: Lista wartości pierwszego wiersza (nagłówka) lub None
        col_idx: Indeks kolumny do sprawdzenia
    
    Returns:
        True jeśli nagłówek kolumny zawiera słowo z blacklisty, False w przeciwnym razie.
        Returns False if header_row is None, col_idx is invalid, or out of bounds.
    """
    if header_row is None:
        return False
    # Handle None col_idx explicitly before numeric comparisons
    if col_idx is None:
        return False
    if col_idx < 0 or col_idx >= len(header_row):
        return False
    
    header_val = header_row[col_idx]
    if header_val is None:
        return False
    
    header_lower = str(header_val).lower().strip()
    
    for blacklisted in COLUMN_BLACKLIST:
        if blacklisted in header_lower:
            return True
    
    return False


def list_spreadsheets_owned_by_me(drive_service, page_size: int = 1000) -> List[Dict[str, Any]]:
    """
    Zwraca listę plików typu spreadsheet, które należą do aktualnego użytkownika.
    """
    files = []
    q = "mimeType='application/vnd.google-apps.spreadsheet' and 'me' in owners"
    page_token = None
    while True:
        resp = (
            drive_service.files()
            .list(q=q, spaces="drive", fields="nextPageToken, files(id, name)", pageSize=page_size, pageToken=page_token)
            .execute()
        )
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def search_in_spreadsheets(
    drive_service,
    sheets_service,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = False,
    max_files: Optional[int] = None,
    stop_event: Optional[threading.Event] = None,
) -> Generator[Dict[str, Any], None, None]:
    """
    Przeszukuje wszystkie arkusze należące do użytkownika wg pattern.
    Zwraca generator wyników:
    {
      "spreadsheetId": ...,
      "spreadsheetName": ...,
      "sheetName": ...,
      "cell": "A1",
      "value": "..."
    }

    Poprawka: jeśli standardowe dopasowanie (substring / regex) nie znajdzie nic,
    a zarówno pattern jak i komórka zawierają cyfry, wykonujemy dopasowanie na
    znormalizowanych ciągach liczbowych.
    """
    files = list_spreadsheets_owned_by_me(drive_service)
    if max_files:
        files = files[:max_files]

    flags = 0 if case_sensitive else re.IGNORECASE
    matcher = re.compile(pattern, flags) if regex else None

    # Pre-compute pattern normalization and check once (optimization)
    pattern_str = pattern if pattern else ""
    pattern_has_digits = bool(re.search(r"\d", pattern_str))
    norm_pat = normalize_number_string(pattern_str) if pattern_has_digits else ""
    digit_pattern = re.compile(r"\d")  # Pre-compiled regex for digit detection

    for f in files:
        # Check stop_event before processing each file
        if stop_event is not None and stop_event.is_set():
            return
        sid = f["id"]
        sname = f.get("name", "")
        # pobierz metadane arkusza (nazwy zakładek)
        try:
            meta = sheets_service.spreadsheets().get(spreadsheetId=sid, fields="sheets.properties").execute()
        except Exception:
            # pomiń nieosiągalne arkusze
            continue
        sheets = meta.get("sheets", [])
        for sh in sheets:
            # Check stop_event before processing each sheet
            if stop_event is not None and stop_event.is_set():
                return
            title = sh["properties"]["title"]
            # odczytaj wszystkie wartości z zakładki (range = title)
            try:
                resp = sheets_service.spreadsheets().values().get(spreadsheetId=sid, range=title, majorDimension="ROWS").execute()
                values = resp.get("values", [])
            except Exception:
                continue
            for r_idx, row in enumerate(values):
                # Check stop_event periodically during row iteration
                if stop_event is not None and stop_event.is_set():
                    return
                if row is None:
                    continue
                for c_idx, cell in enumerate(row):
                    try:
                        # Obsługa None i konwersja do str
                        if cell is None:
                            cell_text = ""
                        elif isinstance(cell, (int, float)):
                            cell_text = str(cell)
                        else:
                            cell_text = str(cell)

                        matched = False
                        # 1) regex match jeśli wybrano regex
                        if regex:
                            try:
                                if matcher and matcher.search(cell_text):
                                    matched = True
                            except re.error:
                                # błędne regex -> pomiń
                                matched = False
                        else:
                            # 2) zwykły substring (case-sensitive lub nie)
                            if pattern and cell_text:
                                if case_sensitive:
                                    if pattern in cell_text:
                                        matched = True
                                else:
                                    if pattern.lower() in cell_text.lower():
                                        matched = True

                        # 3) Jeśli nie znaleziono i pattern i cell zawierają cyfry, spróbuj dopasowania
                        #    po normalizacji liczb (usuń separatory tysięcy, NBSP itp.)
                        if not matched and pattern_has_digits:
                            if digit_pattern.search(cell_text):
                                norm_cell = normalize_number_string(cell_text)
                                if norm_pat and norm_pat in norm_cell:
                                    matched = True

                        if matched:
                            yield {
                                "spreadsheetId": sid,
                                "spreadsheetName": sname,
                                "sheetName": title,
                                "cell": cell_address(r_idx, c_idx),
                                "value": cell_text,
                            }
                    except Exception as e:
                        # Loguj błąd w pojedynczej komórce i kontynuuj wyszukiwanie
                        logger.warning(
                            f"Błąd przetwarzania komórki [{sname}] {title}!{cell_address(r_idx, c_idx)}: {e}"
                        )
                        continue


def search_in_spreadsheet(
    drive_service,
    sheets_service,
    spreadsheet_id: str,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = False,
    search_column_name: Optional[str] = None,
    stop_event: Optional[threading.Event] = None,
    ignore_patterns: Optional[List[str]] = None,
    header_row_indices: Optional[List[int]] = None,
) -> Generator[Dict[str, Any], None, None]:
    """
    Przeszukuje wszystkie zakładki w konkretnym arkuszu wg pattern.
    
    Args:
        drive_service: Obiekt serwisu Google Drive API
        sheets_service: Obiekt serwisu Google Sheets API
        spreadsheet_id: ID arkusza kalkulacyjnego
        pattern: Wzorzec do wyszukania
        regex: Czy użyć wyrażenia regularnego
        case_sensitive: Czy rozróżniać wielkość liter
        search_column_name: Nazwa kolumny do przeszukania lub 'ALL'/'Wszystkie'
        stop_event: Opcjonalny obiekt threading.Event do sygnalizowania zatrzymania
        ignore_patterns: Opcjonalna lista wzorców ignorowania (z parse_ignore_patterns)
        header_row_indices: Opcjonalna lista indeksów wierszy nagłówkowych (0-based)
    
    Zwraca generator wyników w formacie:
    {
      "spreadsheetId": ...,
      "spreadsheetName": ...,
      "sheetName": ...,
      "cell": "A1",
      "searchedValue": "...",
      "stawka": "..."
    }

    Wykorzystuje tę samą normalizację liczb co search_in_spreadsheets.
    Odporność na None/nieoczekiwane typy, przechwytuje błędy per-komórka.
    """
    # Check stop_event at the start
    if stop_event is not None and stop_event.is_set():
        return
    
    # Pobierz nazwę arkusza z metadanych
    spreadsheet_name = ""
    try:
        meta = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id, fields="properties.title,sheets.properties"
        ).execute()
        spreadsheet_name = meta.get("properties", {}).get("title", "")
        sheets = meta.get("sheets", [])
    except Exception as e:
        logger.error(f"Błąd pobierania metadanych arkusza [{spreadsheet_id}]: {e}")
        return

    # Przeszukaj każdą zakładkę
    for sh in sheets:
        # Check stop_event before processing each sheet
        if stop_event is not None and stop_event.is_set():
            return
        sheet_name = sh["properties"]["title"]
        yield from search_in_sheet(
            drive_service,
            sheets_service,
            spreadsheet_id=spreadsheet_id,
            sheet_name=sheet_name,
            pattern=pattern,
            regex=regex,
            case_sensitive=case_sensitive,
            search_column_name=search_column_name,
            spreadsheet_name=spreadsheet_name,
            stop_event=stop_event,
            ignore_patterns=ignore_patterns,
            header_row_indices=header_row_indices,
        )


def search_in_sheet(
    drive_service,
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = False,
    search_column_name: Optional[str] = None,
    spreadsheet_name: Optional[str] = None,
    stop_event: Optional[threading.Event] = None,
    ignore_patterns: Optional[List[str]] = None,
    header_row_indices: Optional[List[int]] = None,
) -> Generator[Dict[str, Any], None, None]:
    """
    Przeszukuje tylko wybraną zakładkę w konkretnym arkuszu wg pattern.
    
    Args:
        drive_service: Obiekt serwisu Google Drive API
        sheets_service: Obiekt serwisu Google Sheets API
        spreadsheet_id: ID arkusza kalkulacyjnego
        sheet_name: Nazwa zakładki
        pattern: Wzorzec do wyszukania
        regex: Czy użyć wyrażenia regularnego
        case_sensitive: Czy rozróżniać wielkość liter
        search_column_name: Nazwa kolumny do przeszukania, 'ALL'/'Wszystkie' lub None
            - Jeśli 'ALL'/'Wszystkie' - przeszukuj wszystkie kolumny
            - Jeśli konkretna nazwa - przeszukuj tylko tę kolumnę
            - Jeśli None - tryb strict: przeszukuj tylko kolumnę 'numer zlecenia'
        spreadsheet_name: Opcjonalna nazwa arkusza (unika dodatkowego wywołania API)
        stop_event: Opcjonalny obiekt threading.Event do sygnalizowania zatrzymania
        ignore_patterns: Opcjonalna lista wzorców ignorowania (z parse_ignore_patterns)
            - Kolumny pasujące do wzorców są pomijane nawet jeśli pasują do search_column_name
        header_row_indices: Opcjonalna lista indeksów wierszy nagłówkowych (0-based)
            - Jeśli podana, używa tych wierszy do budowy połączonych nagłówków
            - Jeśli None, wykrywa automatycznie (wiersz 1 lub 2)
    
    Zwraca generator wyników w formacie:
    {
      "spreadsheetId": ...,
      "spreadsheetName": ...,
      "sheetName": ...,
      "cell": "A1",
      "searchedValue": "..." (wartość z przeszukanej kolumny),
      "stawka": "..." (wartość z kolumny Stawka)
    }

    Logika wykrywania nagłówków:
    - Jeśli header_row_indices podane, używa tych wierszy i łączy wartości spacją
    - W przeciwnym razie sprawdza wiersz 1, czy zawiera szukaną kolumnę
    - Jeśli nie znajdzie w wierszu 1, sprawdza wiersz 2
    - Kolumna 'Stawka' jest pobierana z tego samego wiersza nagłówków
    - Dane są przeszukiwane od wiersza następującego po nagłówku

    Logika wyszukiwania:
    1. Jeśli search_column_name == 'ALL'/'Wszystkie':
       - Przeszukuj wszystkie kolumny
       - Zawsze pobieraj stawkę z kolumny rozpoznanej jako 'Stawka'
    2. Jeśli search_column_name jest podane i istnieje:
       - Przeszukuj tylko tę kolumnę
       - Pobieraj stawkę z kolumny 'Stawka'
    3. Jeśli search_column_name nie jest podane (None):
       - Tryb strict: przeszukuj tylko kolumnę 'numer zlecenia'
       - Jeśli jej brak, nie zwracaj wyników

    Wykorzystuje normalizację liczb dla porównań numerycznych.
    Obsługuje URL-e (wyciąga numeric tokens).
    Odporność na None/nieoczekiwane typy, przechwytuje błędy per-komórka.
    """
    # Check stop_event at the start
    if stop_event is not None and stop_event.is_set():
        return
    
    # Użyj przekazanej nazwy arkusza lub pobierz ją z API
    if spreadsheet_name is None:
        try:
            meta = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id, fields="properties.title"
            ).execute()
            spreadsheet_name = meta.get("properties", {}).get("title", "")
        except Exception as e:
            logger.warning(f"Nie można pobrać nazwy arkusza [{spreadsheet_id}]: {e}")
            spreadsheet_name = spreadsheet_id

    flags = 0 if case_sensitive else re.IGNORECASE
    matcher = re.compile(pattern, flags) if regex else None

    # Pre-compute pattern normalization and check once (optimization)
    pattern_str = pattern if pattern else ""
    pattern_has_digits = bool(re.search(r"\d", pattern_str))
    norm_pat = normalize_number_string(pattern_str) if pattern_has_digits else ""
    digit_pattern = re.compile(r"\d")  # Pre-compiled regex for digit detection

    # Pobierz wartości z wybranej zakładki
    try:
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            majorDimension="ROWS"
        ).execute()
        values = resp.get("values", [])
    except Exception as e:
        logger.error(f"Błąd pobierania danych z arkusza [{spreadsheet_name}] {sheet_name}: {e}")
        return

    if not values:
        return

    # Wykryj wiersz nagłówków (może być w wierszu 1 lub 2, lub wiele wierszy jeśli header_row_indices podane)
    # Przekazujemy search_column_name i header_row_indices do detect_header_row
    header_row_idx, header_row, start_row = detect_header_row(values, search_column_name, header_row_indices)
    
    # Znajdź kolumnę stawki (w tym samym wierszu nagłówków)
    stawka_idx = find_stawka_column_index(header_row) if header_row else None
    
    # Określ tryb wyszukiwania
    search_all = is_search_all_columns(search_column_name)
    target_col_indices = []
    
    if not search_all and search_column_name is not None:
        # Szukamy konkretnej kolumny - znajdź WSZYSTKIE kolumny o tej nazwie (z filtrowaniem ignorowanych)
        target_col_indices = find_all_column_indices_by_name(header_row, search_column_name, ignore_patterns) if header_row else []
        if not target_col_indices:
            # Kolumna nie istnieje lub wszystkie są ignorowane
            if header_row and find_all_column_indices_by_name(header_row, search_column_name, None):
                logger.debug(f"Wszystkie kolumny '{search_column_name}' są ignorowane w [{spreadsheet_name}] {sheet_name}")
            else:
                logger.debug(f"Kolumna '{search_column_name}' nie istnieje w [{spreadsheet_name}] {sheet_name}")
            return
    elif search_column_name is None:
        # Tryb strict - szukaj tylko 'numer zlecenia'
        zlecenie_idx, _ = find_header_indices(header_row) if header_row else (None, None)
        if zlecenie_idx is None:
            # Brak kolumny 'numer zlecenia' - nie zwracaj wyników
            logger.debug(f"Brak kolumny 'numer zlecenia' w [{spreadsheet_name}] {sheet_name}")
            return
        # Sprawdź czy kolumna 'numer zlecenia' nie jest ignorowana
        if header_row and ignore_patterns and matches_ignore_pattern(str(header_row[zlecenie_idx]), ignore_patterns):
            logger.debug(f"Kolumna 'numer zlecenia' jest ignorowana w [{spreadsheet_name}] {sheet_name}")
            return
        target_col_indices = [zlecenie_idx]

    def check_match(cell_text: str) -> bool:
        """Sprawdza czy komórka pasuje do wzorca."""
        matched = False
        # 1) regex match jeśli wybrano regex
        if regex:
            try:
                if matcher and matcher.search(cell_text):
                    matched = True
            except re.error:
                matched = False
        else:
            # 2) zwykły substring (case-sensitive lub nie)
            if pattern and cell_text:
                if case_sensitive:
                    if pattern in cell_text:
                        matched = True
                else:
                    if pattern.lower() in cell_text.lower():
                        matched = True

        # 3) Jeśli nie znaleziono i pattern i cell zawierają cyfry, spróbuj dopasowania
        #    po normalizacji liczb (usuń separatory tysięcy, NBSP itp.)
        if not matched and pattern_has_digits:
            if digit_pattern.search(cell_text):
                norm_cell = normalize_number_string(cell_text)
                if norm_pat and norm_pat in norm_cell:
                    matched = True
        
        # 4) Dla URL-ów: wyciągnij tokeny numeryczne i sprawdź
        if not matched and pattern_has_digits:
            cell_text_lower = cell_text.lower()
            if 'http://' in cell_text_lower or 'https://' in cell_text_lower or 'www.' in cell_text_lower:
                numeric_tokens = extract_numeric_tokens(cell_text)
                for token in numeric_tokens:
                    norm_token = normalize_number_string(token)
                    if norm_pat and norm_pat in norm_token:
                        matched = True
                        break
        
        return matched

    def get_stawka_for_row(row: List[Any], match_col_idx: int) -> str:
        """Pobiera wartość stawki dla wiersza."""
        if stawka_idx is not None:
            return get_cell_value_safe(row, stawka_idx) or ""
        else:
            # Fallback: wartość z komórki po prawej (jeśli nie ma na blackliście)
            next_col_idx = match_col_idx + 1
            if is_column_blacklisted(header_row, next_col_idx):
                return ""
            return get_cell_value_safe(row, next_col_idx) or ""

    if search_all:
        # Tryb 'ALL' - przeszukuj wszystkie kolumny (z pominięciem ignorowanych)
        for r_idx in range(start_row, len(values)):
            # Check stop_event periodically during row iteration
            if stop_event is not None and stop_event.is_set():
                return
            row = values[r_idx]
            if row is None:
                continue
            for c_idx, cell in enumerate(row):
                try:
                    # Sprawdź czy kolumna nie jest ignorowana
                    if header_row and c_idx < len(header_row) and ignore_patterns:
                        if matches_ignore_pattern(str(header_row[c_idx]), ignore_patterns):
                            continue  # Pomiń ignorowane kolumny
                    
                    # Obsługa None i konwersja do str
                    if cell is None:
                        cell_text = ""
                    elif isinstance(cell, (int, float)):
                        cell_text = str(cell)
                    else:
                        cell_text = str(cell)

                    if check_match(cell_text):
                        # Sprawdź czy wartość komórki nie jest ignorowana
                        if ignore_patterns and matches_ignore_value(cell_text, ignore_patterns):
                            continue  # Pomiń ignorowane wartości
                        
                        stawka_value = get_stawka_for_row(row, c_idx)
                        
                        yield {
                            "spreadsheetId": spreadsheet_id,
                            "spreadsheetName": spreadsheet_name,
                            "sheetName": sheet_name,
                            "cell": cell_address(r_idx, c_idx),
                            "searchedValue": cell_text,
                            "stawka": stawka_value,
                        }
                except Exception as e:
                    logger.warning(
                        f"Błąd przetwarzania komórki [{spreadsheet_name}] {sheet_name}!{cell_address(r_idx, c_idx)}: {e}"
                    )
                    continue
    else:
        # Iterate through all matching columns
        for r_idx in range(start_row, len(values)):
            # Check stop_event periodically during row iteration
            if stop_event is not None and stop_event.is_set():
                return
            row = values[r_idx]
            if row is None:
                continue
            try:
                # Iteruj przez wszystkie kolumny pasujące do nazwy
                for target_col_idx in target_col_indices:
                    # Pobierz wartość z docelowej kolumny
                    cell_value = get_cell_value_safe(row, target_col_idx)
                    if cell_value is None:
                        continue
                    
                    if check_match(cell_value):
                        # Sprawdź czy wartość komórki nie jest ignorowana
                        if ignore_patterns and matches_ignore_value(cell_value, ignore_patterns):
                            continue  # Pomiń ignorowane wartości
                        
                        stawka_value = get_stawka_for_row(row, target_col_idx)
                        
                        yield {
                            "spreadsheetId": spreadsheet_id,
                            "spreadsheetName": spreadsheet_name,
                            "sheetName": sheet_name,
                            "cell": cell_address(r_idx, target_col_idx),
                            "searchedValue": cell_value,
                            "stawka": stawka_value,
                        }
            except Exception as e:
                logger.warning(
                    f"Błąd przetwarzania wiersza [{spreadsheet_name}] {sheet_name}!{r_idx+1}: {e}"
                )


def search_across_spreadsheets(
    drive_service,
    sheets_service,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = False,
    search_column_name: Optional[str] = None,
    spreadsheet_ids: Optional[List[str]] = None,
    stop_event: Optional[threading.Event] = None,
    ignore_patterns: Optional[List[str]] = None,
    header_row_indices: Optional[List[int]] = None,
) -> Generator[Dict[str, Any], None, None]:
    """
    Przeszukuje wiele arkuszy kalkulacyjnych wg pattern.
    
    Gdy spreadsheet_ids jest None, pobiera wszystkie arkusze użytkownika
    przez list_spreadsheets_owned_by_me i iteruje po nich.
    Gdy spreadsheet_ids jest podane, iteruje tylko po podanych ID.
    
    Dla każdego arkusza wywołuje search_in_spreadsheet i yielduje wyniki.
    Błąd przy jednym arkuszu nie przerywa całego procesu.
    
    Args:
        drive_service: Obiekt serwisu Google Drive API
        sheets_service: Obiekt serwisu Google Sheets API
        pattern: Wzorzec do wyszukania
        regex: Czy użyć wyrażenia regularnego
        case_sensitive: Czy rozróżniać wielkość liter
        search_column_name: Nazwa kolumny do przeszukania lub 'ALL'/'Wszystkie'
        spreadsheet_ids: Lista ID arkuszy do przeszukania lub None (wszystkie)
        stop_event: Opcjonalny obiekt threading.Event do sygnalizowania zatrzymania
        ignore_patterns: Opcjonalna lista wzorców ignorowania (z parse_ignore_patterns)
        header_row_indices: Opcjonalna lista indeksów wierszy nagłówkowych (0-based)
    
    Yields:
        Wyniki w formacie:
        {
          "spreadsheetId": ...,
          "spreadsheetName": ...,
          "sheetName": ...,
          "cell": "A1",
          "searchedValue": "...",
          "stawka": "..."
        }
    """
    # Check stop_event at the start
    if stop_event is not None and stop_event.is_set():
        return
    
    # Pobierz listę arkuszy do przeszukania
    if spreadsheet_ids is None:
        try:
            files = list_spreadsheets_owned_by_me(drive_service)
            spreadsheet_list = [(f["id"], f.get("name", "")) for f in files]
        except Exception as e:
            logger.error(f"Błąd pobierania listy arkuszy: {e}")
            return
    else:
        spreadsheet_list = [(sid, "") for sid in spreadsheet_ids]
    
    # Iteruj po wszystkich arkuszach
    for spreadsheet_id, spreadsheet_name in spreadsheet_list:
        # Check stop_event before processing each spreadsheet
        if stop_event is not None and stop_event.is_set():
            return
        try:
            results_gen = search_in_spreadsheet(
                drive_service,
                sheets_service,
                spreadsheet_id=spreadsheet_id,
                pattern=pattern,
                regex=regex,
                case_sensitive=case_sensitive,
                search_column_name=search_column_name,
                stop_event=stop_event,
                ignore_patterns=ignore_patterns,
                header_row_indices=header_row_indices,
            )
            for result in results_gen:
                # Check stop_event after each result
                if stop_event is not None and stop_event.is_set():
                    return
                yield result
        except Exception as e:
            # Błąd przy jednym arkuszu nie przerywa całego procesu
            logger.warning(f"Błąd przeszukiwania arkusza [{spreadsheet_id}]: {e}")
            continue


def find_duplicates_in_sheet(
    drive_service,
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    search_column_name: str,
    normalize: bool = True,
    spreadsheet_name: Optional[str] = None,
    stop_event: Optional[threading.Event] = None,
) -> List[Dict[str, Any]]:
    """
    Wykrywa duplikaty wartości w wskazanej kolumnie arkusza.
    
    Args:
        drive_service: Obiekt serwisu Google Drive API
        sheets_service: Obiekt serwisu Google Sheets API
        spreadsheet_id: ID arkusza kalkulacyjnego
        sheet_name: Nazwa zakładki
        search_column_name: Nazwa kolumny do analizy duplikatów
        normalize: Czy normalizować wartości (strip, lowercase, normalize_number_string)
        spreadsheet_name: Opcjonalna nazwa arkusza (unika dodatkowego wywołania API)
        stop_event: Opcjonalny obiekt threading.Event do sygnalizowania zatrzymania
    
    Returns:
        Lista obiektów reprezentujących znalezione duplikaty:
        {
            "spreadsheetId": str,
            "spreadsheetName": str,
            "sheetName": str,
            "columnName": str,
            "value": str,
            "count": int,
            "rows": [int],  # 1-based row indices
            "sample_cells": [str]  # raw cell values
        }
        Pusta lista jeśli kolumna nie istnieje lub brak duplikatów.
    """
    # Check stop_event at the start
    if stop_event is not None and stop_event.is_set():
        return []
    
    # Użyj przekazanej nazwy arkusza lub pobierz ją z API
    if spreadsheet_name is None:
        try:
            meta = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id, fields="properties.title"
            ).execute()
            spreadsheet_name = meta.get("properties", {}).get("title", "")
        except Exception as e:
            logger.warning(f"Nie można pobrać nazwy arkusza [{spreadsheet_id}]: {e}")
            spreadsheet_name = spreadsheet_id
    
    # Pobierz wartości z wybranej zakładki
    try:
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            majorDimension="ROWS"
        ).execute()
        values = resp.get("values", [])
    except Exception as e:
        logger.error(f"Błąd pobierania danych z arkusza [{spreadsheet_name}] {sheet_name}: {e}")
        return []
    
    if not values:
        return []
    
    # Wykryj wiersz nagłówków
    header_row_idx, header_row, start_row = detect_header_row(values, search_column_name)
    
    if header_row is None:
        logger.debug(f"Brak wiersza nagłówków w [{spreadsheet_name}] {sheet_name}")
        return []
    
    # Znajdź wszystkie indeksy kolumn pasujących do nazwy
    target_col_indices = find_all_column_indices_by_name(header_row, search_column_name)
    
    if not target_col_indices:
        logger.debug(f"Kolumna '{search_column_name}' nie istnieje w [{spreadsheet_name}] {sheet_name}")
        return []
    
    # Detect duplicates separately in each matching column
    all_duplicates = []
    
    for target_col_idx in target_col_indices:
        # Map normalized values to their occurrences: normalized_value -> [(row_index_1based, raw_value), ...]
        value_occurrences: Dict[str, List[Tuple[int, str]]] = {}
        
        # Iteruj przez wiersze danych dla tej kolumny
        for r_idx in range(start_row, len(values)):
            # Check stop_event periodically
            if stop_event is not None and stop_event.is_set():
                return []
            
            row = values[r_idx]
            if row is None:
                continue
            
            try:
                cell_value = get_cell_value_safe(row, target_col_idx)
                if cell_value is None or cell_value.strip() == "":
                    continue
                
                raw_value = cell_value
                
                # Normalizuj wartość
                if normalize:
                    # Dla liczb użyj normalize_number_string
                    normalized = normalize_number_string(cell_value)
                    if not normalized:
                        # Dla tekstu: strip + lowercase
                        normalized = cell_value.strip().lower()
                else:
                    normalized = cell_value
                
                # 1-based row index (API zwraca 0-based, ale wyświetlamy 1-based)
                row_1based = r_idx + 1
                
                if normalized not in value_occurrences:
                    value_occurrences[normalized] = []
                value_occurrences[normalized].append((row_1based, raw_value))
                
            except Exception as e:
                logger.warning(
                    f"Błąd przetwarzania wiersza [{spreadsheet_name}] {sheet_name}!{r_idx+1}: {e}"
                )
                continue
        
        # Filtruj tylko duplikaty (count > 1) dla tej kolumny
        for normalized_value, occurrences in value_occurrences.items():
            if len(occurrences) > 1:
                rows = [occ[0] for occ in occurrences]
                sample_cells = [occ[1] for occ in occurrences[:5]]  # Max 5 przykładów
                
                # Użyj oryginalnej wartości z pierwszego wystąpienia
                display_value = occurrences[0][1]
                
                # Dodaj informację o kolumnie (A1 notation) do nazwy kolumny jeśli jest wiele kolumn
                column_display_name = search_column_name
                if len(target_col_indices) > 1:
                    column_display_name = f"{search_column_name} (kolumna {col_index_to_a1(target_col_idx)})"
                
                all_duplicates.append({
                    "spreadsheetId": spreadsheet_id,
                    "spreadsheetName": spreadsheet_name,
                    "sheetName": sheet_name,
                    "columnName": column_display_name,
                    "value": display_value,
                    "count": len(occurrences),
                    "rows": rows,
                    "sample_cells": sample_cells,
                })
    
    return all_duplicates


def find_duplicates_across_spreadsheets(
    drive_service,
    sheets_service,
    spreadsheet_ids: Optional[List[str]],
    search_column_name: str,
    normalize: bool = True,
    stop_event: Optional[threading.Event] = None,
) -> Generator[Dict[str, Any], None, None]:
    """
    Wykrywa duplikaty wartości w wskazanej kolumnie we wielu arkuszach.
    
    Iteruje po podanych arkuszach (lub wszystkich arkuszach użytkownika jeśli
    spreadsheet_ids jest None) i używa find_duplicates_in_sheet dla każdego.
    
    Args:
        drive_service: Obiekt serwisu Google Drive API
        sheets_service: Obiekt serwisu Google Sheets API
        spreadsheet_ids: Lista ID arkuszy do przeszukania lub None (wszystkie)
        search_column_name: Nazwa kolumny do analizy duplikatów
        normalize: Czy normalizować wartości
        stop_event: Opcjonalny obiekt threading.Event do sygnalizowania zatrzymania
    
    Yields:
        Wyniki w formacie:
        {
            "spreadsheetId": str,
            "spreadsheetName": str,
            "sheetName": str,
            "columnName": str,
            "value": str,
            "count": int,
            "rows": [int],
            "sample_cells": [str]
        }
    """
    # Check stop_event at the start
    if stop_event is not None and stop_event.is_set():
        return
    
    # Pobierz listę arkuszy do przeszukania
    if spreadsheet_ids is None:
        try:
            files = list_spreadsheets_owned_by_me(drive_service)
            spreadsheet_list = [(f["id"], f.get("name", "")) for f in files]
        except Exception as e:
            logger.error(f"Błąd pobierania listy arkuszy: {e}")
            return
    else:
        spreadsheet_list = [(sid, "") for sid in spreadsheet_ids]
    
    # Iteruj po wszystkich arkuszach
    for spreadsheet_id, spreadsheet_name in spreadsheet_list:
        # Check stop_event before processing each spreadsheet
        if stop_event is not None and stop_event.is_set():
            return
        
        try:
            # Pobierz metadane arkusza (nazwy zakładek)
            meta = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id, fields="properties.title,sheets.properties"
            ).execute()
            
            if not spreadsheet_name:
                spreadsheet_name = meta.get("properties", {}).get("title", spreadsheet_id)
            
            sheets = meta.get("sheets", [])
            
            # Iteruj po wszystkich zakładkach
            for sh in sheets:
                # Check stop_event before processing each sheet
                if stop_event is not None and stop_event.is_set():
                    return
                
                sheet_name = sh["properties"]["title"]
                
                # Znajdź duplikaty w tej zakładce
                duplicates = find_duplicates_in_sheet(
                    drive_service,
                    sheets_service,
                    spreadsheet_id=spreadsheet_id,
                    sheet_name=sheet_name,
                    search_column_name=search_column_name,
                    normalize=normalize,
                    spreadsheet_name=spreadsheet_name,
                    stop_event=stop_event,
                )
                
                # Yield każdy duplikat
                for dup in duplicates:
                    if stop_event is not None and stop_event.is_set():
                        return
                    yield dup
                    
        except Exception as e:
            logger.warning(f"Błąd przetwarzania arkusza [{spreadsheet_id}]: {e}")
            continue
