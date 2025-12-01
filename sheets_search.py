"""
sheets_search.py
Funkcje:
- list_spreadsheets_owned_by_me(drive_service)
- search_in_spreadsheets(drive_service, sheets_service, pattern, regex=False, case_sensitive=False, max_files=None)
- search_in_sheet(sheets_service, spreadsheet_id, spreadsheet_name, sheet_name, pattern, regex=False, case_sensitive=False)
- search_in_spreadsheet(drive_service, sheets_service, spreadsheet_id, pattern, regex=False, case_sensitive=False)

Poprawka: normalizacja ciągów liczbowych, aby wyszukiwanie znajdowało liczby pomimo różnego formatowania (spacje, NBSP, separatory tysięcy, przecinek/kropka).
Dodatkowa odporność na wartości None i typy numeryczne (int/float).

Nowa funkcjonalność:
- Wykrywanie nagłówków "Numer zlecenia" i "Stawka" w arkuszu
- Zwracanie searchedValue (numer zlecenia) i stawka w wynikach
- Fallback: jeśli brak nagłówków, stawka z komórki po prawej
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
                if variant in cell_lower or cell_lower == variant:
                    zlecenie_idx = idx
                    break
        
        # Sprawdź warianty nagłówka "Stawka"
        if stawka_idx is None:
            for variant in STAWKA_HEADERS:
                if variant in cell_lower or cell_lower == variant:
                    stawka_idx = idx
                    break
    
    return zlecenie_idx, stawka_idx


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
) -> Generator[Dict[str, Any], None, None]:
    """
    Przeszukuje wszystkie zakładki w konkretnym arkuszu wg pattern.
    
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
            sheets_service,
            spreadsheet_id=spreadsheet_id,
            spreadsheet_name=spreadsheet_name,
            sheet_name=sheet_name,
            pattern=pattern,
            regex=regex,
            case_sensitive=case_sensitive,
        )


def search_in_sheet(
    sheets_service,
    spreadsheet_id: str,
    spreadsheet_name: str,
    sheet_name: str,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = False,
) -> Generator[Dict[str, Any], None, None]:
    """
    Przeszukuje tylko wybraną zakładkę w konkretnym arkuszu wg pattern.
    
    Zwraca generator wyników w formacie:
    {
      "spreadsheetId": ...,
      "spreadsheetName": ...,
      "sheetName": ...,
      "cell": "A1",
      "searchedValue": "..." (wartość z kolumny Numer zlecenia lub znaleziona komórka),
      "stawka": "..." (wartość z kolumny Stawka lub komórki po prawej)
    }

    Logika wyszukiwania:
    1. Jeśli pierwszy wiersz zawiera nagłówki "Numer zlecenia" i "Stawka":
       - Przeszukuj tylko kolumnę "Numer zlecenia"
       - Zwróć odpowiadającą wartość z kolumny "Stawka"
    2. Fallback (brak nagłówków):
       - Przeszukuj wszystkie komórki
       - Zwróć wartość z komórki po prawej jako stawka

    Wykorzystuje normalizację liczb dla porównań numerycznych.
    Odporność na None/nieoczekiwane typy, przechwytuje błędy per-komórka.
    """
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
    zlecenie_idx, stawka_idx = None, None
    header_mode = False
    start_row = 0
    
    if is_likely_header_row(first_row):
        zlecenie_idx, stawka_idx = find_header_indices(first_row)
        if zlecenie_idx is not None and stawka_idx is not None:
            header_mode = True
            start_row = 1  # Pomiń wiersz nagłówka
            logger.debug(f"Wykryto nagłówki w [{spreadsheet_name}] {sheet_name}: "
                        f"Zlecenie={zlecenie_idx}, Stawka={stawka_idx}")

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
        
        return matched

    if header_mode:
        # Tryb z nagłówkami - przeszukuj tylko kolumnę "Numer zlecenia"
        for r_idx in range(start_row, len(values)):
            row = values[r_idx]
            if row is None:
                continue
            try:
                # Pobierz wartość z kolumny zlecenia
                zlecenie_value = get_cell_value_safe(row, zlecenie_idx)
                if zlecenie_value is None:
                    continue
                
                if check_match(zlecenie_value):
                    # Pobierz wartość stawki
                    stawka_value = get_cell_value_safe(row, stawka_idx) or ""
                    
                    yield {
                        "spreadsheetId": spreadsheet_id,
                        "spreadsheetName": spreadsheet_name,
                        "sheetName": sheet_name,
                        "cell": cell_address(r_idx, zlecenie_idx),
                        "searchedValue": zlecenie_value,
                        "stawka": stawka_value,
                    }
            except Exception as e:
                logger.warning(
                    f"Błąd przetwarzania wiersza [{spreadsheet_name}] {sheet_name}!{r_idx+1}: {e}"
                )
                continue
    else:
        # Tryb fallback - przeszukuj wszystkie komórki
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

                    if check_match(cell_text):
                        # Fallback: stawka to wartość w komórce po prawej
                        stawka_value = get_cell_value_safe(row, c_idx + 1) or ""
                        
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
