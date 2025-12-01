"""
sheets_search.py
Funkcje:
- list_spreadsheets_owned_by_me(drive_service)
- search_in_spreadsheets(drive_service, sheets_service, pattern, regex=False, case_sensitive=False, max_files=None)
- search_in_sheet(drive_service, sheets_service, spreadsheet_id, sheet_name, pattern, regex=False, case_sensitive=False, search_column_name=None)
- search_in_spreadsheet(drive_service, sheets_service, spreadsheet_id, pattern, regex=False, case_sensitive=False, search_column_name=None)
- get_sheet_headers(sheets_service, spreadsheet_id, sheet_name)

Poprawka: normalizacja ciągów liczbowych, aby wyszukiwanie znajdowało liczby pomimo różnego formatowania (spacje, NBSP, separatory tysięcy, przecinek/kropka).
Dodatkowa odporność na wartości None i typy numeryczne (int/float).

Nowa funkcjonalność:
- Wykrywanie nagłówków "Numer zlecenia" i "Stawka" w arkuszu
- Zwracanie searchedValue (numer zlecenia) i stawka w wynikach
- Fallback: jeśli brak nagłówków, stawka z komórki po prawej
- Obsługa parametru search_column_name: 'ALL'/'Wszystkie' przeszukuje wszystkie kolumny,
  konkretna nazwa kolumny przeszukuje tylko tę kolumnę, brak wartości przeszukuje 'numer zlecenia'
"""

import logging
import re
from typing import List, Dict, Any, Generator, Optional, Union

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


def get_sheet_headers(sheets_service, spreadsheet_id: str, sheet_name: str) -> List[str]:
    """
    Pobiera nagłówki (pierwszy wiersz) z arkusza.
    
    Args:
        sheets_service: Obiekt serwisu Google Sheets API
        spreadsheet_id: ID arkusza kalkulacyjnego
        sheet_name: Nazwa zakładki
    
    Returns:
        Lista nagłówków kolumn (puste stringi dla pustych komórek)
    """
    try:
        # Pobierz tylko pierwszy wiersz
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!1:1",
            majorDimension="ROWS"
        ).execute()
        values = resp.get("values", [])
        if values and len(values) > 0:
            # Konwertuj wszystkie wartości do stringów
            return [str(cell) if cell is not None else "" for cell in values[0]]
        return []
    except Exception as e:
        logger.error(f"Błąd pobierania nagłówków z [{spreadsheet_id}] {sheet_name}: {e}")
        return []


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
            title = sh["properties"]["title"]
            # odczytaj wszystkie wartości z zakładki (range = title)
            try:
                resp = sheets_service.spreadsheets().values().get(spreadsheetId=sid, range=title, majorDimension="ROWS").execute()
                values = resp.get("values", [])
            except Exception:
                continue
            for r_idx, row in enumerate(values):
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
    
    Zwraca generator wyników w formacie:
    {
      "spreadsheetId": ...,
      "spreadsheetName": ...,
      "sheetName": ...,
      "cell": "A1",
      "searchedValue": "..." (wartość z przeszukanej kolumny),
      "stawka": "..." (wartość z kolumny Stawka)
    }

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

    # Sprawdź czy pierwszy wiersz to nagłówek
    first_row = values[0] if values else []
    has_header = is_likely_header_row(first_row)
    start_row = 1 if has_header else 0
    header_row = first_row if has_header else None
    
    # Znajdź kolumnę stawki
    stawka_idx = find_stawka_column_index(first_row) if has_header else None
    
    # Określ tryb wyszukiwania
    search_all = is_search_all_columns(search_column_name)
    target_col_idx = None
    
    if not search_all and search_column_name is not None:
        # Szukamy konkretnej kolumny
        target_col_idx = find_column_index_by_name(first_row, search_column_name)
        if target_col_idx is None:
            # Kolumna nie istnieje - nie zwracaj wyników dla tej zakładki
            logger.debug(f"Kolumna '{search_column_name}' nie istnieje w [{spreadsheet_name}] {sheet_name}")
            return
    elif search_column_name is None:
        # Tryb strict - szukaj tylko 'numer zlecenia'
        zlecenie_idx, _ = find_header_indices(first_row)
        if zlecenie_idx is None:
            # Brak kolumny 'numer zlecenia' - nie zwracaj wyników
            logger.debug(f"Brak kolumny 'numer zlecenia' w [{spreadsheet_name}] {sheet_name}")
            return
        target_col_idx = zlecenie_idx

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
        # Tryb 'ALL' - przeszukuj wszystkie kolumny
        for r_idx in range(start_row, len(values)):
            row = values[r_idx]
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

                    if check_match(cell_text):
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
        # Tryb konkretnej kolumny (target_col_idx jest ustawiony)
        for r_idx in range(start_row, len(values)):
            row = values[r_idx]
            if row is None:
                continue
            try:
                # Pobierz wartość z docelowej kolumny
                cell_value = get_cell_value_safe(row, target_col_idx)
                if cell_value is None:
                    continue
                
                if check_match(cell_value):
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
