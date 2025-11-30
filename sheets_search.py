"""
sheets_search.py
Funkcje:
- list_spreadsheets_owned_by_me(drive_service)
- search_in_spreadsheets(drive_service, sheets_service, pattern, regex=False, case_sensitive=False, max_files=None)

Poprawka: normalizacja ciągów liczbowych, aby wyszukiwanie znajdowało liczby pomimo różnego formatowania (spacje, NBSP, separatory tysięcy, przecinek/kropka).
"""

import re
from typing import List, Dict, Any, Generator, Optional

# helpers
def col_index_to_a1(n: int) -> str:
    """Konwertuje indeks kolumny 0-based na etykietę A1 (0 -> A)."""
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_address(row_idx: int, col_idx: int) -> str:
    return f"{col_index_to_a1(col_idx)}{row_idx + 1}"


def normalize_number_string(s: str) -> str:
    """
    Normalizuje ciąg zawierający liczbę:
    - usuwa spacje zwykłe i NBSP, znaki cienkiej spacji itp.
    - zamienia przecinek na kropkę (dla miejsc dziesiętnych)
    - usuwa wszystko poza cyframi, kropką i minusem
    Zwraca znormalizowany ciąg, lub '' jeśli po usunięciu nic nie zostaje.
    """
    if s is None:
        return ""
    # usuń niełamiące spacje i inne biały znaki grupujące
    s = s.replace("\u00A0", "").replace("\u202F", "")
    # usuń zwykłe spacje
    s = s.replace(" ", "")
    # zamień przecinek na kropkę (np. "1,23" -> "1.23")
    s = s.replace(",", ".")
    # zostaw tylko cyfry, kropkę i minus
    s = re.sub(r"[^\d\.\-]", "", s)
    return s


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
                for c_idx, cell in enumerate(row):
                    cell_text = str(cell)
                    matched = False
                    # 1) regex match jeśli wybrano regex
                    if regex:
                        try:
                            if matcher.search(cell_text):
                                matched = True
                        except re.error:
                            # błędne regex -> pomiń
                            matched = False
                    else:
                        # 2) zwykły substring (case-sensitive lub nie)
                        if case_sensitive:
                            if pattern in cell_text:
                                matched = True
                        else:
                            if pattern.lower() in cell_text.lower():
                                matched = True

                    # 3) Jeśli nie znaleziono i pattern i cell zawierają cyfry, spróbuj dopasowania
                    #    po normalizacji liczb (usuń separatory tysięcy, NBSP itp.)
                    if not matched:
                        if re.search(r"\d", pattern or "") and re.search(r"\d", cell_text or ""):
                            norm_cell = normalize_number_string(cell_text)
                            norm_pat = normalize_number_string(pattern or "")
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
