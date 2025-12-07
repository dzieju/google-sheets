"""
gui.py
GUI oparty na FreeSimpleGUI do przeszukiwania Google Sheets.
Wykorzystuje istniejące moduły google_auth i sheets_search.
"""

import json
import os
import re
import threading

import FreeSimpleGUI as sg

# Import existing modules
from google_auth import build_services, TOKEN_FILE
from sheets_search import (
    list_spreadsheets_owned_by_me,
    search_in_spreadsheets,
    search_in_sheet,
    search_in_spreadsheet,
    search_across_spreadsheets,
    find_duplicates_in_sheet,
    find_duplicates_across_spreadsheets,
    ALL_COLUMNS_VALUES,
    parse_ignore_patterns,
    parse_header_rows,
)
from quadra_service import (
    read_dbf_column,
    search_dbf_values_in_sheets,
    format_quadra_result_for_table,
    export_quadra_results_to_json,
    export_quadra_results_to_csv,
    get_dbf_field_names,
    read_dbf_records_with_extra_fields,
    detect_dbf_field_name,
    DBF_NUMER_FIELD_NAMES,
    DBF_STAWKA_FIELD_NAMES,
    DBF_CZESCI_FIELD_NAMES,
)


# -------------------- Settings persistence --------------------
# Path to settings file for persistent app configuration
SETTINGS_FILE = os.path.join(os.path.expanduser('~'), '.google_sheets_settings.json')


def load_settings():
    """
    Load application settings from JSON file.
    Returns empty dict if file doesn't exist or is corrupted.
    Used to restore user preferences like DBF mappings and last used paths.
    """
    if not os.path.exists(SETTINGS_FILE):
        return {}
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        # If file is corrupted or unreadable, return empty dict
        return {}


def save_settings(settings):
    """
    Save application settings to JSON file.
    Persists user preferences like DBF field mappings and last used paths
    so they are remembered between application runs.
    """
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except OSError:
        # If we can't save settings, silently fail (non-critical)
        pass


# -------------------- Events --------------------
EVENT_AUTH_DONE = "-AUTH_DONE-"
EVENT_FILES_LOADED = "-FILES_LOADED-"
EVENT_SHEETS_LOADED = "-SHEETS_LOADED-"
EVENT_PREVIEW_LOADED = "-PREVIEW_LOADED-"
EVENT_SEARCH_RESULT = "-SEARCH_RESULT-"
EVENT_SEARCH_DONE = "-SEARCH_DONE-"
EVENT_ERROR = "-ERROR-"
# Events for single sheet search
EVENT_SS_FILES_LOADED = "-SS_FILES_LOADED-"
EVENT_SS_SHEETS_LOADED = "-SS_SHEETS_LOADED-"
EVENT_SS_SEARCH_RESULT = "-SS_SEARCH_RESULT-"
EVENT_SS_SEARCH_DONE = "-SS_SEARCH_DONE-"
# Events for duplicate detection
EVENT_DUP_RESULT = "-DUP_RESULT-"
EVENT_DUP_DONE = "-DUP_DONE-"
# Events for Quadra tab
EVENT_QUADRA_FILES_LOADED = "-QUADRA_FILES_LOADED-"
EVENT_QUADRA_SHEETS_LOADED = "-QUADRA_SHEETS_LOADED-"
EVENT_QUADRA_CHECK_DONE = "-QUADRA_CHECK_DONE-"

# -------------------- Global state --------------------
drive_service = None
sheets_service = None
current_spreadsheets = []
search_thread = None
stop_search_flag = threading.Event()
# Global state for single sheet search
ss_current_spreadsheets = []
ss_current_sheets = []
ss_search_thread = None
ss_stop_search_flag = threading.Event()
# Global state for duplicate detection
dup_search_thread = None
dup_stop_search_flag = threading.Event()
# Global state for Quadra tab
quadra_current_spreadsheets = []
quadra_current_sheets = []
quadra_check_thread = None
quadra_stop_flag = threading.Event()
quadra_dbf_field_mapping = {}  # Stores user-configured DBF field mapping
quadra_dbf_field_names = []  # Stores available DBF field names


# -------------------- Helper functions --------------------
def format_result(result: dict) -> str:
    """Format a single search result for display (legacy format for main search tab)."""
    return (
        f"[{result['spreadsheetName']}] "
        f"Arkusz: {result['sheetName']}, "
        f"Komórka: {result['cell']}, "
        f"Wartość: {result.get('value', result.get('searchedValue', ''))}"
    )


def format_ss_result_for_table(result: dict) -> list:
    """Format a single sheet search result as table row [Arkusz, Arkusz kalkulacyjny, Zlecenie, Stawka]."""
    return [
        result.get('sheetName', ''),  # Arkusz (tab/sheet name)
        result.get('spreadsheetName', ''),  # Arkusz kalkulacyjny
        result.get('searchedValue', ''),  # Zlecenie
        result.get('stawka', ''),  # Stawka
    ]


def format_dup_result_for_table(result: dict) -> list:
    """Format a duplicate result as table row [Arkusz, Kolumna, Wartość, Ile razy, Przykładowe wiersze]."""
    rows = result.get('rows', [])
    # Show first 5 row numbers as example
    example_rows = ', '.join(str(r) for r in rows[:5])
    if len(rows) > 5:
        example_rows += f'... (+{len(rows) - 5})'
    
    return [
        result.get('sheetName', ''),  # Arkusz (tab/sheet name)
        result.get('columnName', ''),
        result.get('value', ''),
        str(result.get('count', 0)),
        example_rows,
    ]


# -------------------- Background thread functions --------------------
def authenticate_thread(window):
    """Run OAuth authentication in background thread."""
    global drive_service, sheets_service
    try:
        drive_service, sheets_service = build_services()
        window.write_event_value(EVENT_AUTH_DONE, "success")
    except FileNotFoundError as e:
        window.write_event_value(EVENT_ERROR, str(e))
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd autoryzacji: {e}")


def load_files_thread(window):
    """Load spreadsheets list in background thread."""
    global current_spreadsheets
    try:
        if drive_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        files = list_spreadsheets_owned_by_me(drive_service)
        current_spreadsheets = files
        window.write_event_value(EVENT_FILES_LOADED, files)
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd ładowania plików: {e}")


def load_sheets_for_file_thread(window, spreadsheet_id, spreadsheet_name):
    """Load sheet names for a selected spreadsheet."""
    try:
        if sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        meta = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id, fields="sheets.properties"
        ).execute()
        sheet_names = [sh["properties"]["title"] for sh in meta.get("sheets", [])]
        window.write_event_value(EVENT_SHEETS_LOADED, {"id": spreadsheet_id, "name": spreadsheet_name, "sheets": sheet_names})
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd ładowania arkuszy: {e}")


def load_preview_thread(window, spreadsheet_id, sheet_name):
    """Load preview of a sheet (first 20 rows)."""
    try:
        if sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            majorDimension="ROWS"
        ).execute()
        values = resp.get("values", [])
        # Show first 20 rows as preview
        preview_lines = []
        for row in values[:20]:
            preview_lines.append("\t".join(str(c) for c in row))
        if len(values) > 20:
            preview_lines.append(f"... (jeszcze {len(values) - 20} wierszy)")
        window.write_event_value(EVENT_PREVIEW_LOADED, "\n".join(preview_lines))
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd ładowania podglądu: {e}")


def search_thread_func(window, pattern, regex, case_sensitive, max_files):
    """Run search in background thread, posting results to GUI."""
    global stop_search_flag
    try:
        if drive_service is None or sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return

        stop_search_flag.clear()
        results_gen = search_in_spreadsheets(
            drive_service,
            sheets_service,
            pattern=pattern,
            regex=regex,
            case_sensitive=case_sensitive,
            max_files=max_files if max_files > 0 else None,
            stop_event=stop_search_flag,
        )

        for result in results_gen:
            if stop_search_flag.is_set():
                window.write_event_value(EVENT_SEARCH_DONE, "stopped")
                return
            window.write_event_value(EVENT_SEARCH_RESULT, result)

        window.write_event_value(EVENT_SEARCH_DONE, "completed")

    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd wyszukiwania: {e}")
        window.write_event_value(EVENT_SEARCH_DONE, "error")


# -------------------- Single Sheet Search Thread Functions --------------------
def ss_load_files_thread(window):
    """Load spreadsheets list for single sheet search tab."""
    global ss_current_spreadsheets
    try:
        if drive_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        files = list_spreadsheets_owned_by_me(drive_service)
        ss_current_spreadsheets = files
        window.write_event_value(EVENT_SS_FILES_LOADED, files)
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd ładowania plików: {e}")


def ss_load_sheets_thread(window, spreadsheet_id, spreadsheet_name):
    """Load sheet names for selected spreadsheet in single sheet search tab."""
    global ss_current_sheets
    try:
        if sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        meta = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id, fields="sheets.properties"
        ).execute()
        sheet_names = [sh["properties"]["title"] for sh in meta.get("sheets", [])]
        ss_current_sheets = sheet_names
        window.write_event_value(EVENT_SS_SHEETS_LOADED, {
            "id": spreadsheet_id,
            "name": spreadsheet_name,
            "sheets": sheet_names
        })
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd ładowania arkuszy: {e}")


def ss_search_thread_func(window, spreadsheet_id, spreadsheet_name, sheet_name, pattern, regex, case_sensitive, all_sheets=False, search_column_name=None, ignore_patterns=None, header_row_indices=None):
    """Run search in a single sheet or all sheets in background thread."""
    global ss_stop_search_flag
    try:
        if sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return

        ss_stop_search_flag.clear()
        
        if all_sheets:
            # Search in all sheets of the spreadsheet
            results_gen = search_in_spreadsheet(
                drive_service,
                sheets_service,
                spreadsheet_id=spreadsheet_id,
                pattern=pattern,
                regex=regex,
                case_sensitive=case_sensitive,
                search_column_name=search_column_name,
                stop_event=ss_stop_search_flag,
                ignore_patterns=ignore_patterns,
                header_row_indices=header_row_indices,
            )
        else:
            # Search in a single sheet
            results_gen = search_in_sheet(
                drive_service,
                sheets_service,
                spreadsheet_id=spreadsheet_id,
                sheet_name=sheet_name,
                pattern=pattern,
                regex=regex,
                case_sensitive=case_sensitive,
                search_column_name=search_column_name,
                stop_event=ss_stop_search_flag,
                ignore_patterns=ignore_patterns,
                header_row_indices=header_row_indices,
            )

        for result in results_gen:
            if ss_stop_search_flag.is_set():
                window.write_event_value(EVENT_SS_SEARCH_DONE, "stopped")
                return
            window.write_event_value(EVENT_SS_SEARCH_RESULT, result)

        window.write_event_value(EVENT_SS_SEARCH_DONE, "completed")

    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd wyszukiwania: {e}")
        window.write_event_value(EVENT_SS_SEARCH_DONE, "error")


def ss_search_all_spreadsheets_thread_func(window, pattern, regex, case_sensitive, search_column_name=None, ignore_patterns=None, header_row_indices=None):
    """Run search across all user's spreadsheets in background thread."""
    global ss_stop_search_flag
    try:
        if drive_service is None or sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return

        ss_stop_search_flag.clear()
        
        # Search across all spreadsheets using the new backend function
        results_gen = search_across_spreadsheets(
            drive_service,
            sheets_service,
            pattern=pattern,
            regex=regex,
            case_sensitive=case_sensitive,
            search_column_name=search_column_name,
            spreadsheet_ids=None,  # None means search all spreadsheets
            stop_event=ss_stop_search_flag,
            ignore_patterns=ignore_patterns,
            header_row_indices=header_row_indices,
        )

        for result in results_gen:
            if ss_stop_search_flag.is_set():
                window.write_event_value(EVENT_SS_SEARCH_DONE, "stopped")
                return
            window.write_event_value(EVENT_SS_SEARCH_RESULT, result)

        window.write_event_value(EVENT_SS_SEARCH_DONE, "completed")

    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd wyszukiwania: {e}")
        window.write_event_value(EVENT_SS_SEARCH_DONE, "error")


# -------------------- Duplicate Detection Thread Functions --------------------
def dup_search_thread_func(window, spreadsheet_id, spreadsheet_name, sheet_name, search_column_name, all_sheets=False):
    """Run duplicate detection in a single sheet or all sheets in background thread."""
    global dup_stop_search_flag
    try:
        if sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return

        dup_stop_search_flag.clear()
        
        if all_sheets:
            # Get all sheets in the spreadsheet
            try:
                meta = sheets_service.spreadsheets().get(
                    spreadsheetId=spreadsheet_id, fields="sheets.properties"
                ).execute()
                sheets = meta.get("sheets", [])
            except Exception as e:
                window.write_event_value(EVENT_ERROR, f"Błąd pobierania zakładek: {e}")
                window.write_event_value(EVENT_DUP_DONE, "error")
                return
            
            for sh in sheets:
                if dup_stop_search_flag.is_set():
                    window.write_event_value(EVENT_DUP_DONE, "stopped")
                    return
                
                sheet_title = sh["properties"]["title"]
                duplicates = find_duplicates_in_sheet(
                    drive_service,
                    sheets_service,
                    spreadsheet_id=spreadsheet_id,
                    sheet_name=sheet_title,
                    search_column_name=search_column_name,
                    normalize=True,
                    spreadsheet_name=spreadsheet_name,
                    stop_event=dup_stop_search_flag,
                )
                
                for dup in duplicates:
                    if dup_stop_search_flag.is_set():
                        window.write_event_value(EVENT_DUP_DONE, "stopped")
                        return
                    window.write_event_value(EVENT_DUP_RESULT, dup)
        else:
            # Search in a single sheet
            duplicates = find_duplicates_in_sheet(
                drive_service,
                sheets_service,
                spreadsheet_id=spreadsheet_id,
                sheet_name=sheet_name,
                search_column_name=search_column_name,
                normalize=True,
                spreadsheet_name=spreadsheet_name,
                stop_event=dup_stop_search_flag,
            )
            
            for dup in duplicates:
                if dup_stop_search_flag.is_set():
                    window.write_event_value(EVENT_DUP_DONE, "stopped")
                    return
                window.write_event_value(EVENT_DUP_RESULT, dup)

        window.write_event_value(EVENT_DUP_DONE, "completed")

    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd wykrywania duplikatów: {e}")
        window.write_event_value(EVENT_DUP_DONE, "error")


def dup_search_all_spreadsheets_thread_func(window, search_column_name):
    """Run duplicate detection across all user's spreadsheets in background thread."""
    global dup_stop_search_flag
    try:
        if drive_service is None or sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return

        dup_stop_search_flag.clear()
        
        # Find duplicates across all spreadsheets
        results_gen = find_duplicates_across_spreadsheets(
            drive_service,
            sheets_service,
            spreadsheet_ids=None,  # None means all spreadsheets
            search_column_name=search_column_name,
            normalize=True,
            stop_event=dup_stop_search_flag,
        )

        for result in results_gen:
            if dup_stop_search_flag.is_set():
                window.write_event_value(EVENT_DUP_DONE, "stopped")
                return
            window.write_event_value(EVENT_DUP_RESULT, result)

        window.write_event_value(EVENT_DUP_DONE, "completed")

    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd wykrywania duplikatów: {e}")
        window.write_event_value(EVENT_DUP_DONE, "error")


# -------------------- Quadra Thread Functions --------------------
def quadra_load_files_thread(window):
    """Load spreadsheets list for Quadra tab."""
    global quadra_current_spreadsheets
    try:
        if drive_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        files = list_spreadsheets_owned_by_me(drive_service)
        quadra_current_spreadsheets = files
        window.write_event_value(EVENT_QUADRA_FILES_LOADED, files)
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd ładowania plików: {e}")


def quadra_load_sheets_thread(window, spreadsheet_id, spreadsheet_name):
    """Load sheet names for selected spreadsheet in Quadra tab."""
    global quadra_current_sheets
    try:
        if sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        meta = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id, fields="sheets.properties"
        ).execute()
        sheet_names = [sh["properties"]["title"] for sh in meta.get("sheets", [])]
        quadra_current_sheets = sheet_names
        window.write_event_value(EVENT_QUADRA_SHEETS_LOADED, {
            "id": spreadsheet_id,
            "name": spreadsheet_name,
            "sheets": sheet_names
        })
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd ładowania arkuszy: {e}")


def quadra_check_thread_func(window, dbf_path, dbf_column, spreadsheet_id, mode, sheet_names, column_names, mapping=None):
    """Run Quadra check in background thread."""
    global quadra_stop_flag
    try:
        if sheets_service is None:
            window.write_event_value(EVENT_ERROR, "Najpierw zaloguj się.")
            return
        
        quadra_stop_flag.clear()
        
        # Read DBF file with extra fields (stawka, czesci)
        try:
            from quadra_service import read_dbf_records_with_extra_fields
            dbf_records = read_dbf_records_with_extra_fields(dbf_path, dbf_column, mapping)
            if not dbf_records:
                window.write_event_value(EVENT_ERROR, "Brak wartości w wybranej kolumnie DBF.")
                window.write_event_value(EVENT_QUADRA_CHECK_DONE, "error")
                return
        except Exception as e:
            window.write_event_value(EVENT_ERROR, f"Błąd odczytu pliku DBF: {e}")
            window.write_event_value(EVENT_QUADRA_CHECK_DONE, "error")
            return
        
        # Search in Google Sheets
        try:
            results = search_dbf_values_in_sheets(
                drive_service=drive_service,
                sheets_service=sheets_service,
                dbf_values=dbf_records,
                spreadsheet_id=spreadsheet_id,
                mode=mode,
                sheet_names=sheet_names,
                column_names=column_names,
                header_row_index=0
            )
            window.write_event_value(EVENT_QUADRA_CHECK_DONE, results)
        except Exception as e:
            window.write_event_value(EVENT_ERROR, f"Błąd przeszukiwania arkuszy: {e}")
            window.write_event_value(EVENT_QUADRA_CHECK_DONE, "error")
    
    except Exception as e:
        window.write_event_value(EVENT_ERROR, f"Błąd sprawdzania Quadra: {e}")
        window.write_event_value(EVENT_QUADRA_CHECK_DONE, "error")


# -------------------- GUI Layout --------------------
def create_auth_tab():
    """Create Authorization tab layout."""
    return [
        [sg.Text("Status:", size=(10, 1)), sg.Text("Nie zalogowano", key="-AUTH_STATUS-", size=(40, 1))],
        [sg.Button("Zaloguj się", key="-AUTH_BTN-"), sg.Button("Wyczyść token", key="-CLEAR_TOKEN-")],
        [sg.HorizontalSeparator()],
        [sg.Text("Aby się zalogować, kliknij przycisk 'Zaloguj się'.", font=("Helvetica", 9))],
        [sg.Text("Otworzy się przeglądarka do autoryzacji OAuth.", font=("Helvetica", 9))],
        [sg.Text("Aby wymusić ponowne logowanie, kliknij 'Wyczyść token'.", font=("Helvetica", 9))],
    ]


def create_files_tab():
    """Create Files & Preview tab layout."""
    return [
        [sg.Button("Odśwież listę plików", key="-REFRESH_FILES-")],
        [sg.Text("Twoje arkusze:")],
        [sg.Listbox(values=[], size=(50, 10), key="-FILES_LIST-", enable_events=True)],
        [sg.Text("Arkusze (zakładki) w wybranym pliku:")],
        [sg.Listbox(values=[], size=(50, 5), key="-SHEETS_LIST-", enable_events=True)],
        [sg.Text("Podgląd arkusza:")],
        [sg.Multiline(size=(70, 10), key="-PREVIEW-", disabled=True)],
    ]


def create_search_tab():
    """Create Search tab layout."""
    return [
        [sg.Text("Zapytanie:"), sg.Input(key="-SEARCH_QUERY-", size=(40, 1))],
        [sg.Checkbox("Regex", key="-REGEX-"), sg.Checkbox("Rozróżniaj wielkość liter", key="-CASE_SENSITIVE-")],
        [sg.Text("Maks. plików:"), sg.Input(key="-MAX_FILES-", size=(10, 1), default_text="")],
        [sg.Button("Szukaj", key="-SEARCH_START-"), sg.Button("Zatrzymaj", key="-SEARCH_STOP-", disabled=True)],
        [sg.HorizontalSeparator()],
        [sg.Text("Wyniki:", font=("Helvetica", 10, "bold"))],
        [sg.Multiline(size=(80, 15), key="-SEARCH_RESULTS-", disabled=True, autoscroll=True)],
        [sg.Text("Znaleziono: 0", key="-SEARCH_COUNT-")],
        [sg.Button("Wyczyść wyniki", key="-CLEAR_RESULTS-"), sg.Button("Zapisz do JSON", key="-SAVE_JSON-")],
    ]


def create_settings_tab():
    """Create Settings tab layout."""
    return [
        [sg.Text("Ustawienia aplikacji", font=("Helvetica", 12, "bold"))],
        [sg.HorizontalSeparator()],
        [sg.Text("Ścieżka do token.json:"), sg.Text(TOKEN_FILE, key="-TOKEN_PATH-")],
        [sg.Text("Token istnieje:"), sg.Text("Nie", key="-TOKEN_EXISTS-")],
        [sg.HorizontalSeparator()],
        [sg.Text("Informacje o aplikacji:", font=("Helvetica", 10, "bold"))],
        [sg.Text("Aplikacja do przeszukiwania arkuszy Google Sheets.", font=("Helvetica", 9))],
        [sg.Text("Wykorzystuje Google Drive API i Google Sheets API.", font=("Helvetica", 9))],
    ]


def create_single_sheet_search_tab():
    """Create Single Sheet Search tab layout."""
    # Definicja kolumn tabeli wyników (dodano "Arkusz" jako pierwszą kolumnę)
    table_headings = ["Arkusz", "Arkusz kalkulacyjny", "Zlecenie", "Stawka"]
    # Definicja kolumn tabeli duplikatów (dodano "Arkusz" jako pierwszą kolumnę)
    dup_table_headings = ["Arkusz", "Kolumna", "Wartość", "Ile razy", "Przykładowe wiersze"]
    
    return [
        [sg.Text("Przeszukiwanie pojedynczego arkusza", font=("Helvetica", 12, "bold"))],
        [sg.HorizontalSeparator()],
        [sg.Button("Odśwież listę arkuszy", key="-SS_REFRESH_FILES-")],
        [sg.Text("Wybierz arkusz:"), sg.Checkbox("Wybierz wszystkie arkusze", key="-SSPREADSHEETS_SELECT_ALL-", enable_events=True)],
        [sg.Combo(values=[], key="-SSPREADSHEETS_DROPDOWN-", enable_events=True, readonly=True, expand_x=True)],
        [sg.Text("Wybierz zakładkę:"), sg.Checkbox("Wybierz wszystkie", key="-SHEET_ALL_SHEETS-", enable_events=True)],
        [sg.Combo(values=[], key="-SSHEETS_DROPDOWN-", enable_events=True, readonly=True, expand_x=True)],
        [sg.Text("Kolumna do przeszukania (puste lub 'ALL'/'Wszystkie' = wszystkie):")],
        [sg.Input(key="-SHEET_COLUMN_INPUT-", default_text="", expand_x=True)],
        [sg.HorizontalSeparator()],
        [sg.Text("Zapytanie:"), sg.Input(key="-SHEET_QUERY-", expand_x=True)],
        [sg.Text("Ignoruj (puste = brak, oddziel przecinkiem/średnikiem/nową linią, obsługuje wildcards *):")],
        [sg.Multiline(key="-SHEET_IGNORE-", size=(None, 3), default_text="", expand_x=True)],
        [sg.Text("Header rows (domyślnie '1', można podać wiele oddzielone przecinkami np. '1,2'):"), sg.Input(key="-HEADER_ROWS-", default_text="1", size=(10, 1))],
        [sg.Checkbox("Regex", key="-SHEET_REGEX-"), sg.Checkbox("Rozróżniaj wielkość liter", key="-SHEET_CASE-")],
        [
            sg.Button("Szukaj", key="-SHEET_SEARCH_BTN-"),
            sg.Button("Znajdź duplikaty", key="-DUP_SEARCH_BTN-"),
            sg.Button("Zatrzymaj", key="-SHEET_SEARCH_STOP-", disabled=True)
        ],
        [sg.HorizontalSeparator()],
        [sg.Text("Wyniki wyszukiwania:", font=("Helvetica", 10, "bold"))],
        [sg.Table(
            values=[],
            headings=table_headings,
            key="-SHEET_RESULTS_TABLE-",
            auto_size_columns=True,
            justification='left',
            num_rows=10,
            expand_x=True,
            expand_y=False,
            enable_events=False,
            vertical_scroll_only=False,
        )],
        [sg.Text("Znaleziono: 0", key="-SS_SEARCH_COUNT-")],
        [sg.Button("Wyczyść wyniki", key="-SS_CLEAR_RESULTS-"), sg.Button("Zapisz do JSON", key="-SHEET_SAVE_RESULTS-")],
        [sg.HorizontalSeparator()],
        [sg.Text("Wyniki duplikatów:", font=("Helvetica", 10, "bold"))],
        [sg.Table(
            values=[],
            headings=dup_table_headings,
            key="-DUP_RESULTS_TABLE-",
            auto_size_columns=True,
            justification='left',
            num_rows=10,
            expand_x=True,
            expand_y=True,
            enable_events=False,
            vertical_scroll_only=False,
        )],
        [sg.Text("Znaleziono duplikatów: 0", key="-DUP_SEARCH_COUNT-")],
        [sg.Button("Wyczyść duplikaty", key="-DUP_CLEAR_RESULTS-"), sg.Button("Zapisz duplikaty do JSON", key="-DUP_SAVE_RESULTS-")],
    ]


def create_quadra_tab():
    """Create Quadra tab layout for checking DBF order numbers against Google Sheets."""
    # Updated table headings: Arkusz, Numer z DBF, Stawka, Czesci, Status, Kolumna, Wiersz, Uwagi
    table_headings = ["Arkusz", "Numer z DBF", "Stawka", "Czesci", "Status", "Kolumna", "Wiersz", "Uwagi"]
    
    return [
        [sg.Text("Quadra: Sprawdzanie numerów zleceń z DBF", font=("Helvetica", 12, "bold"))],
        [sg.HorizontalSeparator()],
        
        # DBF file selection
        [sg.Text("Plik DBF:", size=(15, 1)), sg.Input(key="-QUADRA_DBF_PATH-", expand_x=True, readonly=True, enable_events=True), 
         sg.FileBrowse("Wybierz plik DBF", key="-QUADRA_DBF_BROWSE-", file_types=(("DBF Files", "*.dbf"), ("All Files", "*.*")))],
        [sg.Text("Kolumna DBF:", size=(15, 1)), sg.Input(key="-QUADRA_DBF_COLUMN-", default_text="B", size=(10, 1)),
         sg.Text("(litera A, B, C... lub numer 1, 2, 3...)"),
         sg.Button("Konfiguruj mapowanie pól", key="-QUADRA_CONFIG_MAPPING-", size=(20, 1))],
        
        # DBF field mapping panel (initially hidden)
        [sg.pin(sg.Column([
            [sg.Text("Mapowanie pól DBF:", font=("Helvetica", 10, "bold"))],
            [sg.Text("Numer z DBF:", size=(15, 1)), sg.Combo(values=[], key="-QUADRA_MAP_NUMER-", readonly=True, size=(20, 1))],
            [sg.Text("Stawka:", size=(15, 1)), sg.Combo(values=[], key="-QUADRA_MAP_STAWKA-", readonly=True, size=(20, 1))],
            [sg.Text("Części:", size=(15, 1)), sg.Combo(values=[], key="-QUADRA_MAP_CZESCI-", readonly=True, size=(20, 1))],
            [sg.Text("Płatnik:", size=(15, 1)), sg.Combo(values=[], key="-QUADRA_MAP_PLATNIK-", readonly=True, size=(20, 1))],
            [sg.Button("Zastosuj mapowanie", key="-QUADRA_APPLY_MAPPING-"), sg.Button("Resetuj", key="-QUADRA_RESET_MAPPING-")],
        ], key="-QUADRA_MAPPING_PANEL-", visible=False))],
        
        [sg.HorizontalSeparator()],
        
        # Spreadsheet selection
        [sg.Button("Odśwież listę arkuszy", key="-QUADRA_REFRESH_FILES-")],
        [sg.Text("Wybierz arkusz kalkulacyjny:")],
        [sg.Combo(values=[], key="-QUADRA_SPREADSHEET_DROPDOWN-", enable_events=True, readonly=True, expand_x=True)],
        [sg.Text("Wybierz zakładki:"), sg.Checkbox("Wszystkie zakładki", key="-QUADRA_ALL_SHEETS-", default=True, enable_events=True)],
        [sg.Combo(values=[], key="-QUADRA_SHEETS_DROPDOWN-", enable_events=True, readonly=True, expand_x=True, disabled=True)],
        
        [sg.HorizontalSeparator()],
        
        # Search options
        [sg.Text("Opcje wyszukiwania:")],
        [sg.Radio("Exact (trim, case-insensitive, numeric)", "QUADRA_MODE", key="-QUADRA_EXACT-", default=True),
         sg.Radio("Substring", "QUADRA_MODE", key="-QUADRA_SUBSTRING-")],
        [sg.Text("Kolumny do przeszukania (puste = wszystkie):")],
        [sg.Input(key="-QUADRA_COLUMN_FILTER-", expand_x=True)],
        
        [sg.HorizontalSeparator()],
        
        # Action buttons
        [sg.Button("Sprawdź", key="-QUADRA_CHECK_BTN-", size=(15, 1)),
         sg.Button("Zatrzymaj", key="-QUADRA_STOP_BTN-", disabled=True, size=(15, 1))],
        
        [sg.HorizontalSeparator()],
        
        # Results table
        [sg.Text("Wyniki:", font=("Helvetica", 10, "bold"))],
        [sg.Table(
            values=[],
            headings=table_headings,
            key="-QUADRA_RESULTS_TABLE-",
            auto_size_columns=True,
            justification='left',
            num_rows=15,
            expand_x=True,
            expand_y=True,
            enable_events=False,
            vertical_scroll_only=False,
        )],
        [sg.Text("Znaleziono: 0 | Brakujących: 0", key="-QUADRA_STATUS-")],
        
        # Export buttons
        [sg.Button("Wyczyść wyniki", key="-QUADRA_CLEAR_RESULTS-"),
         sg.Button("Eksportuj JSON", key="-QUADRA_EXPORT_JSON-"),
         sg.Button("Eksportuj CSV", key="-QUADRA_EXPORT_CSV-")],
    ]


def create_layout():
    """Create main window layout with tabs."""
    tab_auth = sg.Tab("Autoryzacja", create_auth_tab())
    tab_files = sg.Tab("Pliki i podgląd", create_files_tab())
    tab_search = sg.Tab("Przeszukiwanie", create_search_tab())
    tab_single_sheet = sg.Tab("Przeszukiwanie arkusza", create_single_sheet_search_tab(), expand_x=True, expand_y=True)
    tab_quadra = sg.Tab("Quadra", create_quadra_tab(), expand_x=True, expand_y=True)
    tab_settings = sg.Tab("Ustawienia", create_settings_tab())

    layout = [
        [sg.TabGroup(
            [[tab_auth, tab_files, tab_search, tab_single_sheet, tab_quadra, tab_settings]], 
            key="-TABGROUP-", 
            expand_x=True, 
            expand_y=True
        )],
        [sg.StatusBar("Gotowe", key="-STATUS_BAR-", size=(60, 1))],
    ]
    return layout


# -------------------- Main GUI loop --------------------
def main():
    """Main function - runs the GUI event loop."""
    global drive_service, sheets_service, search_thread, ss_search_thread, dup_search_thread
    global quadra_dbf_field_mapping, quadra_dbf_field_names

    sg.theme("SystemDefault")

    window = sg.Window(
        "Google Sheets Search",
        create_layout(),
        finalize=True,
        resizable=True
    )

    # State for search results (for JSON export)
    search_results_list = []
    current_spreadsheet_id = None

    # State for single sheet search
    ss_search_results_list = []
    ss_table_data = []  # Data for the results table [Arkusz, Zlecenie, Stawka]
    ss_current_spreadsheet_id = None
    ss_current_spreadsheet_name = None

    # State for duplicate detection
    dup_results_list = []
    dup_table_data = []  # Data for the duplicates table

    # Load settings and restore last DBF path
    app_settings = load_settings()
    window.metadata = {'_app_settings': app_settings}
    
    # Restore last DBF path if it exists
    last_dbf_path = app_settings.get('quadra_last_dbf_path', '')
    if last_dbf_path and os.path.exists(last_dbf_path):
        window["-QUADRA_DBF_PATH-"].update(value=last_dbf_path)
    
    # Update token status on startup
    window["-TOKEN_EXISTS-"].update("Tak" if os.path.exists(TOKEN_FILE) else "Nie")

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED:
            break

        # -------------------- Authorization tab events --------------------
        if event == "-AUTH_BTN-":
            window["-AUTH_STATUS-"].update("Trwa logowanie...")
            window["-STATUS_BAR-"].update("Trwa autoryzacja OAuth...")
            threading.Thread(target=authenticate_thread, args=(window,), daemon=True).start()

        elif event == "-CLEAR_TOKEN-":
            if os.path.exists(TOKEN_FILE):
                try:
                    os.remove(TOKEN_FILE)
                    drive_service = None
                    sheets_service = None
                    window["-AUTH_STATUS-"].update("Nie zalogowano")
                    window["-TOKEN_EXISTS-"].update("Nie")
                    window["-STATUS_BAR-"].update("Token usunięty. Zaloguj się ponownie.")
                    sg.popup("Token został usunięty.", title="Wyczyść token")
                except Exception as e:
                    sg.popup_error(f"Błąd usuwania tokena: {e}")
            else:
                sg.popup("Token nie istnieje.", title="Wyczyść token")

        elif event == EVENT_AUTH_DONE:
            window["-AUTH_STATUS-"].update("Zalogowano pomyślnie")
            window["-TOKEN_EXISTS-"].update("Tak")
            window["-STATUS_BAR-"].update("Autoryzacja zakończona pomyślnie.")

        # -------------------- Files tab events --------------------
        elif event == "-REFRESH_FILES-":
            if drive_service is None:
                sg.popup_error("Najpierw zaloguj się (zakładka Autoryzacja).")
            else:
                window["-STATUS_BAR-"].update("Ładowanie listy plików...")
                threading.Thread(target=load_files_thread, args=(window,), daemon=True).start()

        elif event == EVENT_FILES_LOADED:
            files = values[EVENT_FILES_LOADED]
            display_list = [f"{f['name']}  ({f['id']})" for f in files]
            window["-FILES_LIST-"].update(display_list)
            window["-STATUS_BAR-"].update(f"Załadowano {len(files)} arkuszy.")

        elif event == "-FILES_LIST-":
            selected = values["-FILES_LIST-"]
            if selected:
                # Get file info using the selected index from the listbox
                try:
                    idx = window["-FILES_LIST-"].get_indexes()[0]
                    file_info = current_spreadsheets[idx]
                    current_spreadsheet_id = file_info["id"]
                    window["-STATUS_BAR-"].update(f"Ładowanie arkuszy dla: {file_info['name']}...")
                    threading.Thread(
                        target=load_sheets_for_file_thread,
                        args=(window, file_info["id"], file_info["name"]),
                        daemon=True
                    ).start()
                except (IndexError, KeyError):
                    pass

        elif event == EVENT_SHEETS_LOADED:
            data = values[EVENT_SHEETS_LOADED]
            window["-SHEETS_LIST-"].update(data["sheets"])
            window["-STATUS_BAR-"].update(f"Załadowano {len(data['sheets'])} arkuszy z: {data['name']}")

        elif event == "-SHEETS_LIST-":
            selected = values["-SHEETS_LIST-"]
            if selected and current_spreadsheet_id:
                sheet_name = selected[0]
                window["-STATUS_BAR-"].update(f"Ładowanie podglądu: {sheet_name}...")
                threading.Thread(
                    target=load_preview_thread,
                    args=(window, current_spreadsheet_id, sheet_name),
                    daemon=True
                ).start()

        elif event == EVENT_PREVIEW_LOADED:
            preview_text = values[EVENT_PREVIEW_LOADED]
            window["-PREVIEW-"].update(preview_text)
            window["-STATUS_BAR-"].update("Podgląd załadowany.")

        # -------------------- Search tab events --------------------
        elif event == "-SEARCH_START-":
            query = values["-SEARCH_QUERY-"].strip()
            if not query:
                sg.popup_error("Wprowadź zapytanie do wyszukania.")
                continue

            if drive_service is None or sheets_service is None:
                sg.popup_error("Najpierw zaloguj się (zakładka Autoryzacja).")
                continue

            # Parse max_files
            max_files_str = values["-MAX_FILES-"].strip()
            max_files = None
            if max_files_str:
                try:
                    max_files = int(max_files_str)
                except ValueError:
                    sg.popup_error("Maks. plików musi być liczbą.")
                    continue

            # Clear previous results
            search_results_list.clear()
            window["-SEARCH_RESULTS-"].update("")
            window["-SEARCH_COUNT-"].update("Znaleziono: 0")

            # Disable start, enable stop
            window["-SEARCH_START-"].update(disabled=True)
            window["-SEARCH_STOP-"].update(disabled=False)
            window["-STATUS_BAR-"].update("Trwa wyszukiwanie...")

            # Start search thread
            search_thread = threading.Thread(
                target=search_thread_func,
                args=(
                    window,
                    query,
                    values["-REGEX-"],
                    values["-CASE_SENSITIVE-"],
                    max_files
                ),
                daemon=True
            )
            search_thread.start()

        elif event == "-SEARCH_STOP-":
            stop_search_flag.set()
            window["-STATUS_BAR-"].update("Zatrzymywanie wyszukiwania...")

        elif event == EVENT_SEARCH_RESULT:
            result = values[EVENT_SEARCH_RESULT]
            search_results_list.append(result)
            # Append new result to multiline using print_to_element for efficiency
            new_line = format_result(result)
            window["-SEARCH_RESULTS-"].print(new_line)
            window["-SEARCH_COUNT-"].update(f"Znaleziono: {len(search_results_list)}")

        elif event == EVENT_SEARCH_DONE:
            status = values[EVENT_SEARCH_DONE]
            window["-SEARCH_START-"].update(disabled=False)
            window["-SEARCH_STOP-"].update(disabled=True)
            if status == "completed":
                window["-STATUS_BAR-"].update(f"Wyszukiwanie zakończone. Znaleziono: {len(search_results_list)}")
            elif status == "stopped":
                window["-STATUS_BAR-"].update(f"Wyszukiwanie zatrzymane. Znaleziono: {len(search_results_list)}")
            else:
                window["-STATUS_BAR-"].update("Wyszukiwanie zakończone z błędem.")

        elif event == "-CLEAR_RESULTS-":
            search_results_list.clear()
            window["-SEARCH_RESULTS-"].update("")
            window["-SEARCH_COUNT-"].update("Znaleziono: 0")
            window["-STATUS_BAR-"].update("Wyniki wyczyszczone.")

        elif event == "-SAVE_JSON-":
            if not search_results_list:
                sg.popup("Brak wyników do zapisania.", title="Zapisz do JSON")
                continue
            filename = sg.popup_get_file(
                "Zapisz wyniki do pliku JSON",
                save_as=True,
                default_extension=".json",
                file_types=(("JSON Files", "*.json"), ("All Files", "*.*")),
            )
            if filename:
                try:
                    with open(filename, "w", encoding="utf-8") as f:
                        json.dump(search_results_list, f, ensure_ascii=False, indent=2)
                    sg.popup(f"Zapisano {len(search_results_list)} wyników do:\n{filename}", title="Zapisano")
                    window["-STATUS_BAR-"].update(f"Wyniki zapisane do: {filename}")
                except Exception as e:
                    sg.popup_error(f"Błąd zapisu: {e}")

        # -------------------- Single Sheet Search tab events --------------------
        elif event == "-SS_REFRESH_FILES-":
            if drive_service is None:
                sg.popup_error("Najpierw zaloguj się (zakładka Autoryzacja).")
            else:
                window["-STATUS_BAR-"].update("Ładowanie listy arkuszy...")
                threading.Thread(target=ss_load_files_thread, args=(window,), daemon=True).start()

        elif event == EVENT_SS_FILES_LOADED:
            files = values[EVENT_SS_FILES_LOADED]
            display_list = [f"{f['name']}  ({f['id']})" for f in files]
            window["-SSPREADSHEETS_DROPDOWN-"].update(values=display_list, value="")
            window["-SSHEETS_DROPDOWN-"].update(values=[], value="")
            window["-STATUS_BAR-"].update(f"Załadowano {len(files)} arkuszy.")

        elif event == "-SSPREADSHEETS_DROPDOWN-":
            selected = values["-SSPREADSHEETS_DROPDOWN-"]
            if selected:
                # Find the index in the list
                try:
                    combo_values = window["-SSPREADSHEETS_DROPDOWN-"].Values
                    idx = combo_values.index(selected)
                    file_info = ss_current_spreadsheets[idx]
                    ss_current_spreadsheet_id = file_info["id"]
                    ss_current_spreadsheet_name = file_info["name"]
                    window["-SSHEETS_DROPDOWN-"].update(values=[], value="")
                    # Reset column input when spreadsheet changes
                    window["-SHEET_COLUMN_INPUT-"].update(value="")
                    window["-STATUS_BAR-"].update(f"Ładowanie zakładek dla: {file_info['name']}...")
                    threading.Thread(
                        target=ss_load_sheets_thread,
                        args=(window, file_info["id"], file_info["name"]),
                        daemon=True
                    ).start()
                except (ValueError, IndexError, KeyError):
                    pass

        elif event == EVENT_SS_SHEETS_LOADED:
            data = values[EVENT_SS_SHEETS_LOADED]
            sheets_list = data["sheets"]
            window["-SSHEETS_DROPDOWN-"].update(values=sheets_list, value=sheets_list[0] if len(sheets_list) > 0 else "")
            window["-STATUS_BAR-"].update(f"Załadowano {len(sheets_list)} zakładek z: {data['name']}")

        elif event == "-SSPREADSHEETS_SELECT_ALL-":
            # Toggle spreadsheet dropdown based on checkbox state
            select_all_spreadsheets = values["-SSPREADSHEETS_SELECT_ALL-"]
            # When 'select all spreadsheets' is checked, the dropdown remains visible but search will use all spreadsheets
            # The sheet selection controls are disabled when searching all spreadsheets
            window["-SSHEETS_DROPDOWN-"].update(disabled=select_all_spreadsheets)
            window["-SHEET_ALL_SHEETS-"].update(disabled=select_all_spreadsheets)
            if select_all_spreadsheets:
                window["-SHEET_ALL_SHEETS-"].update(value=True)  # Force all sheets mode

        elif event == "-SHEET_ALL_SHEETS-":
            # Toggle sheet dropdown based on checkbox state
            # Keep column input editable - user can specify a column name even when searching all sheets
            all_sheets_checked = values["-SHEET_ALL_SHEETS-"]
            window["-SSHEETS_DROPDOWN-"].update(disabled=all_sheets_checked)

        elif event == "-SHEET_SEARCH_BTN-":
            query = values["-SHEET_QUERY-"].strip()
            if not query:
                sg.popup_error("Wprowadź zapytanie do wyszukania.")
                continue

            if sheets_service is None:
                sg.popup_error("Najpierw zaloguj się (zakładka Autoryzacja).")
                continue

            select_all_spreadsheets = values["-SSPREADSHEETS_SELECT_ALL-"]
            selected_spreadsheet = values["-SSPREADSHEETS_DROPDOWN-"]
            selected_sheet = values["-SSHEETS_DROPDOWN-"]
            all_sheets_mode = values["-SHEET_ALL_SHEETS-"]
            column_input_value = values["-SHEET_COLUMN_INPUT-"].strip()
            
            # Parse ignore patterns from the Ignoruj field
            ignore_input = values["-SHEET_IGNORE-"].strip()
            ignore_patterns = parse_ignore_patterns(ignore_input) if ignore_input else None
            
            # Parse header rows from the Header rows field
            header_rows_input = values["-HEADER_ROWS-"].strip()
            header_row_indices = parse_header_rows(header_rows_input)

            # When not selecting all spreadsheets, validate spreadsheet selection
            if not select_all_spreadsheets:
                if not selected_spreadsheet:
                    sg.popup_error("Wybierz arkusz z listy lub zaznacz 'Wybierz wszystkie arkusze'.")
                    continue

                if not all_sheets_mode and not selected_sheet:
                    sg.popup_error("Wybierz zakładkę z listy lub zaznacz 'Wybierz wszystkie'.")
                    continue

            # Determine search_column_name based on input field
            # Empty, 'ALL' or 'Wszystkie' (case-insensitive) means search all columns
            # User can specify a column name even when searching all sheets
            if not column_input_value or column_input_value.lower() in ALL_COLUMNS_VALUES:
                search_column_name = "ALL"
            else:
                # User specified a specific column name
                search_column_name = column_input_value

            # Clear previous results
            ss_search_results_list.clear()
            ss_table_data.clear()
            window["-SHEET_RESULTS_TABLE-"].update(values=[])
            window["-SS_SEARCH_COUNT-"].update("Znaleziono: 0")

            # Disable start, enable stop
            window["-SHEET_SEARCH_BTN-"].update(disabled=True)
            window["-SHEET_SEARCH_STOP-"].update(disabled=False)
            
            if select_all_spreadsheets:
                # Search across all spreadsheets owned by user
                window["-STATUS_BAR-"].update("Trwa wyszukiwanie we wszystkich arkuszach...")
                ss_search_thread = threading.Thread(
                    target=ss_search_all_spreadsheets_thread_func,
                    args=(
                        window,
                        query,
                        values["-SHEET_REGEX-"],
                        values["-SHEET_CASE-"],
                        search_column_name,
                        ignore_patterns,
                        header_row_indices
                    ),
                    daemon=True
                )
            else:
                # Get spreadsheet info for single spreadsheet search
                try:
                    combo_values = window["-SSPREADSHEETS_DROPDOWN-"].Values
                    idx = combo_values.index(selected_spreadsheet)
                    file_info = ss_current_spreadsheets[idx]
                    spreadsheet_id = file_info["id"]
                    spreadsheet_name = file_info["name"]
                except (ValueError, IndexError):
                    sg.popup_error("Błąd: nie można znaleźć wybranego arkusza.")
                    window["-SHEET_SEARCH_BTN-"].update(disabled=False)
                    window["-SHEET_SEARCH_STOP-"].update(disabled=True)
                    continue

                if all_sheets_mode:
                    window["-STATUS_BAR-"].update(f"Trwa wyszukiwanie we wszystkich zakładkach: {spreadsheet_name}...")
                else:
                    window["-STATUS_BAR-"].update(f"Trwa wyszukiwanie w: {spreadsheet_name} / {selected_sheet}...")

                # Start search thread for single spreadsheet
                ss_search_thread = threading.Thread(
                    target=ss_search_thread_func,
                    args=(
                        window,
                        spreadsheet_id,
                        spreadsheet_name,
                        selected_sheet,
                        query,
                        values["-SHEET_REGEX-"],
                        values["-SHEET_CASE-"],
                        all_sheets_mode,
                        search_column_name,
                        ignore_patterns,
                        header_row_indices
                    ),
                    daemon=True
                )
            ss_search_thread.start()

        elif event == "-SHEET_SEARCH_STOP-":
            ss_stop_search_flag.set()
            dup_stop_search_flag.set()
            window["-STATUS_BAR-"].update("Zatrzymywanie wyszukiwania...")

        elif event == EVENT_SS_SEARCH_RESULT:
            result = values[EVENT_SS_SEARCH_RESULT]
            ss_search_results_list.append(result)
            table_row = format_ss_result_for_table(result)
            ss_table_data.append(table_row)
            window["-SHEET_RESULTS_TABLE-"].update(values=ss_table_data)
            window["-SS_SEARCH_COUNT-"].update(f"Znaleziono: {len(ss_search_results_list)}")

        elif event == EVENT_SS_SEARCH_DONE:
            status = values[EVENT_SS_SEARCH_DONE]
            window["-SHEET_SEARCH_BTN-"].update(disabled=False)
            window["-SHEET_SEARCH_STOP-"].update(disabled=True)
            if status == "completed":
                window["-STATUS_BAR-"].update(f"Wyszukiwanie zakończone. Znaleziono: {len(ss_search_results_list)}")
            elif status == "stopped":
                window["-STATUS_BAR-"].update(f"Wyszukiwanie zatrzymane. Znaleziono: {len(ss_search_results_list)}")
            else:
                window["-STATUS_BAR-"].update("Wyszukiwanie zakończone z błędem.")

        elif event == "-SS_CLEAR_RESULTS-":
            ss_search_results_list.clear()
            ss_table_data.clear()
            window["-SHEET_RESULTS_TABLE-"].update(values=[])
            window["-SS_SEARCH_COUNT-"].update("Znaleziono: 0")
            window["-STATUS_BAR-"].update("Wyniki wyczyszczone.")

        elif event == "-SHEET_SAVE_RESULTS-":
            if not ss_search_results_list:
                sg.popup("Brak wyników do zapisania.", title="Zapisz do JSON")
                continue
            filename = sg.popup_get_file(
                "Zapisz wyniki do pliku JSON",
                save_as=True,
                default_extension=".json",
                file_types=(("JSON Files", "*.json"), ("All Files", "*.*")),
            )
            if filename:
                try:
                    with open(filename, "w", encoding="utf-8") as f:
                        # Zapisz każdy wynik jako osobny JSON obiekt w linii (JSONL format)
                        for result in ss_search_results_list:
                            # Format wynikowy: {spreadsheetName, sheetName, cell, searchedValue, stawka}
                            export_obj = {
                                "spreadsheetName": result.get("spreadsheetName", ""),
                                "sheetName": result.get("sheetName", ""),
                                "cell": result.get("cell", ""),
                                "searchedValue": result.get("searchedValue", ""),
                                "stawka": result.get("stawka", ""),
                            }
                            f.write(json.dumps(export_obj, ensure_ascii=False) + "\n")
                    sg.popup(f"Zapisano {len(ss_search_results_list)} wyników do:\n{filename}", title="Zapisano")
                    window["-STATUS_BAR-"].update(f"Wyniki zapisane do: {filename}")
                except Exception as e:
                    sg.popup_error(f"Błąd zapisu: {e}")

        # -------------------- Duplicate Detection Events --------------------
        elif event == "-DUP_SEARCH_BTN-":
            column_input_value = values["-SHEET_COLUMN_INPUT-"].strip()
            if not column_input_value:
                sg.popup_error("Podaj nazwę kolumny do analizy duplikatów.")
                continue

            if sheets_service is None:
                sg.popup_error("Najpierw zaloguj się (zakładka Autoryzacja).")
                continue

            select_all_spreadsheets = values["-SSPREADSHEETS_SELECT_ALL-"]
            selected_spreadsheet = values["-SSPREADSHEETS_DROPDOWN-"]
            all_sheets_mode = values["-SHEET_ALL_SHEETS-"]

            # When not selecting all spreadsheets, validate spreadsheet selection
            if not select_all_spreadsheets:
                if not selected_spreadsheet:
                    sg.popup_error("Wybierz arkusz z listy lub zaznacz 'Wybierz wszystkie arkusze'.")
                    continue

            # Clear previous duplicate results
            dup_results_list.clear()
            dup_table_data.clear()
            window["-DUP_RESULTS_TABLE-"].update(values=[])
            window["-DUP_SEARCH_COUNT-"].update("Znaleziono duplikatów: 0")

            # Disable search buttons, enable stop
            window["-SHEET_SEARCH_BTN-"].update(disabled=True)
            window["-DUP_SEARCH_BTN-"].update(disabled=True)
            window["-SHEET_SEARCH_STOP-"].update(disabled=False)

            if select_all_spreadsheets:
                # Detect duplicates across all spreadsheets
                window["-STATUS_BAR-"].update("Trwa wykrywanie duplikatów we wszystkich arkuszach...")
                dup_search_thread = threading.Thread(
                    target=dup_search_all_spreadsheets_thread_func,
                    args=(
                        window,
                        column_input_value
                    ),
                    daemon=True
                )
            else:
                # Get spreadsheet info for single spreadsheet search
                try:
                    combo_values = window["-SSPREADSHEETS_DROPDOWN-"].Values
                    idx = combo_values.index(selected_spreadsheet)
                    file_info = ss_current_spreadsheets[idx]
                    spreadsheet_id = file_info["id"]
                    spreadsheet_name = file_info["name"]
                except (ValueError, IndexError):
                    sg.popup_error("Błąd: nie można znaleźć wybranego arkusza.")
                    window["-SHEET_SEARCH_BTN-"].update(disabled=False)
                    window["-DUP_SEARCH_BTN-"].update(disabled=False)
                    window["-SHEET_SEARCH_STOP-"].update(disabled=True)
                    continue

                selected_sheet = values["-SSHEETS_DROPDOWN-"]
                
                if all_sheets_mode:
                    window["-STATUS_BAR-"].update(f"Trwa wykrywanie duplikatów we wszystkich zakładkach: {spreadsheet_name}...")
                else:
                    if not selected_sheet:
                        sg.popup_error("Wybierz zakładkę z listy lub zaznacz 'Wybierz wszystkie'.")
                        window["-SHEET_SEARCH_BTN-"].update(disabled=False)
                        window["-DUP_SEARCH_BTN-"].update(disabled=False)
                        window["-SHEET_SEARCH_STOP-"].update(disabled=True)
                        continue
                    window["-STATUS_BAR-"].update(f"Trwa wykrywanie duplikatów w: {spreadsheet_name} / {selected_sheet}...")

                dup_search_thread = threading.Thread(
                    target=dup_search_thread_func,
                    args=(
                        window,
                        spreadsheet_id,
                        spreadsheet_name,
                        selected_sheet,
                        column_input_value,
                        all_sheets_mode
                    ),
                    daemon=True
                )
            dup_search_thread.start()

        elif event == EVENT_DUP_RESULT:
            result = values[EVENT_DUP_RESULT]
            dup_results_list.append(result)
            table_row = format_dup_result_for_table(result)
            dup_table_data.append(table_row)
            window["-DUP_RESULTS_TABLE-"].update(values=dup_table_data)
            window["-DUP_SEARCH_COUNT-"].update(f"Znaleziono duplikatów: {len(dup_results_list)}")

        elif event == EVENT_DUP_DONE:
            status = values[EVENT_DUP_DONE]
            window["-SHEET_SEARCH_BTN-"].update(disabled=False)
            window["-DUP_SEARCH_BTN-"].update(disabled=False)
            window["-SHEET_SEARCH_STOP-"].update(disabled=True)
            if status == "completed":
                window["-STATUS_BAR-"].update(f"Wykrywanie duplikatów zakończone. Znaleziono: {len(dup_results_list)}")
            elif status == "stopped":
                window["-STATUS_BAR-"].update(f"Wykrywanie duplikatów zatrzymane. Znaleziono: {len(dup_results_list)}")
            else:
                window["-STATUS_BAR-"].update("Wykrywanie duplikatów zakończone z błędem.")

        elif event == "-DUP_CLEAR_RESULTS-":
            dup_results_list.clear()
            dup_table_data.clear()
            window["-DUP_RESULTS_TABLE-"].update(values=[])
            window["-DUP_SEARCH_COUNT-"].update("Znaleziono duplikatów: 0")
            window["-STATUS_BAR-"].update("Wyniki duplikatów wyczyszczone.")

        elif event == "-DUP_SAVE_RESULTS-":
            if not dup_results_list:
                sg.popup("Brak duplikatów do zapisania.", title="Zapisz do JSON")
                continue
            filename = sg.popup_get_file(
                "Zapisz duplikaty do pliku JSON",
                save_as=True,
                default_extension=".json",
                file_types=(("JSON Files", "*.json"), ("All Files", "*.*")),
            )
            if filename:
                try:
                    with open(filename, "w", encoding="utf-8") as f:
                        # Zapisz każdy wynik jako osobny JSON obiekt w linii (NDJSON format)
                        for result in dup_results_list:
                            export_obj = {
                                "spreadsheetId": result.get("spreadsheetId", ""),
                                "spreadsheetName": result.get("spreadsheetName", ""),
                                "sheetName": result.get("sheetName", ""),
                                "columnName": result.get("columnName", ""),
                                "value": result.get("value", ""),
                                "count": result.get("count", 0),
                                "rows": result.get("rows", []),
                                "sample_cells": result.get("sample_cells", []),
                            }
                            f.write(json.dumps(export_obj, ensure_ascii=False) + "\n")
                    sg.popup(f"Zapisano {len(dup_results_list)} duplikatów do:\n{filename}", title="Zapisano")
                    window["-STATUS_BAR-"].update(f"Duplikaty zapisane do: {filename}")
                except Exception as e:
                    sg.popup_error(f"Błąd zapisu: {e}")

        # -------------------- Quadra Tab Events --------------------
        elif event == "-QUADRA_REFRESH_FILES-":
            if drive_service is None:
                sg.popup_error("Najpierw zaloguj się (zakładka Autoryzacja).")
            else:
                window["-STATUS_BAR-"].update("Ładowanie listy arkuszy...")
                threading.Thread(target=quadra_load_files_thread, args=(window,), daemon=True).start()

        elif event == EVENT_QUADRA_FILES_LOADED:
            files = values[EVENT_QUADRA_FILES_LOADED]
            display_list = [f"{f['name']}  ({f['id']})" for f in files]
            window["-QUADRA_SPREADSHEET_DROPDOWN-"].update(values=display_list, value="")
            window["-QUADRA_SHEETS_DROPDOWN-"].update(values=[], value="")
            window["-STATUS_BAR-"].update(f"Załadowano {len(files)} arkuszy.")

        elif event == "-QUADRA_SPREADSHEET_DROPDOWN-":
            selected = values["-QUADRA_SPREADSHEET_DROPDOWN-"]
            if selected:
                try:
                    combo_values = window["-QUADRA_SPREADSHEET_DROPDOWN-"].Values
                    idx = combo_values.index(selected)
                    file_info = quadra_current_spreadsheets[idx]
                    window["-QUADRA_SHEETS_DROPDOWN-"].update(values=[], value="")
                    window["-STATUS_BAR-"].update(f"Ładowanie zakładek dla: {file_info['name']}...")
                    threading.Thread(
                        target=quadra_load_sheets_thread,
                        args=(window, file_info["id"], file_info["name"]),
                        daemon=True
                    ).start()
                except (ValueError, IndexError, KeyError):
                    pass

        elif event == EVENT_QUADRA_SHEETS_LOADED:
            data = values[EVENT_QUADRA_SHEETS_LOADED]
            sheets_list = data["sheets"]
            window["-QUADRA_SHEETS_DROPDOWN-"].update(values=sheets_list, value=sheets_list[0] if len(sheets_list) > 0 else "")
            window["-STATUS_BAR-"].update(f"Załadowano {len(sheets_list)} zakładek z: {data['name']}")

        elif event == "-QUADRA_ALL_SHEETS-":
            all_sheets_checked = values["-QUADRA_ALL_SHEETS-"]
            window["-QUADRA_SHEETS_DROPDOWN-"].update(disabled=all_sheets_checked)
        
        elif event == "-QUADRA_DBF_PATH-":
            # When DBF file is selected, load field names and auto-populate mapping dropdowns
            dbf_path = values["-QUADRA_DBF_PATH-"].strip()
            if dbf_path and os.path.exists(dbf_path):
                try:
                    quadra_dbf_field_names = get_dbf_field_names(dbf_path)
                    
                    # Update mapping dropdowns with field names
                    field_options = [''] + quadra_dbf_field_names  # Empty option for "not set"
                    window["-QUADRA_MAP_NUMER-"].update(values=field_options, value='')
                    window["-QUADRA_MAP_STAWKA-"].update(values=field_options, value='')
                    window["-QUADRA_MAP_CZESCI-"].update(values=field_options, value='')
                    window["-QUADRA_MAP_PLATNIK-"].update(values=field_options, value='')
                    
                    # Get saved settings from window metadata
                    app_settings = window.metadata.get('_app_settings', {})
                    saved_mapping = app_settings.get('quadra_dbf_field_mapping', {})
                    
                    # If we have saved mappings, try to restore them if fields exist in current DBF
                    if saved_mapping:
                        numer_field = saved_mapping.get('numer_dbf', '')
                        stawka_field = saved_mapping.get('stawka', '')
                        czesci_field = saved_mapping.get('czesci', '')
                        platnik_field = saved_mapping.get('platnik', '')
                        
                        # Only restore if field exists in current DBF
                        if numer_field in quadra_dbf_field_names:
                            window["-QUADRA_MAP_NUMER-"].update(value=numer_field)
                        else:
                            # Auto-detect if saved mapping doesn't exist
                            numer_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_NUMER_FIELD_NAMES)
                            if numer_field:
                                window["-QUADRA_MAP_NUMER-"].update(value=numer_field)
                        
                        if stawka_field in quadra_dbf_field_names:
                            window["-QUADRA_MAP_STAWKA-"].update(value=stawka_field)
                        else:
                            stawka_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_STAWKA_FIELD_NAMES)
                            if stawka_field:
                                window["-QUADRA_MAP_STAWKA-"].update(value=stawka_field)
                        
                        if czesci_field in quadra_dbf_field_names:
                            window["-QUADRA_MAP_CZESCI-"].update(value=czesci_field)
                        else:
                            czesci_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_CZESCI_FIELD_NAMES)
                            if czesci_field:
                                window["-QUADRA_MAP_CZESCI-"].update(value=czesci_field)
                        
                        if platnik_field in quadra_dbf_field_names:
                            window["-QUADRA_MAP_PLATNIK-"].update(value=platnik_field)
                    else:
                        # Auto-detect and set default values
                        numer_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_NUMER_FIELD_NAMES)
                        stawka_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_STAWKA_FIELD_NAMES)
                        czesci_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_CZESCI_FIELD_NAMES)
                        
                        if numer_field:
                            window["-QUADRA_MAP_NUMER-"].update(value=numer_field)
                        if stawka_field:
                            window["-QUADRA_MAP_STAWKA-"].update(value=stawka_field)
                        if czesci_field:
                            window["-QUADRA_MAP_CZESCI-"].update(value=czesci_field)
                    
                    # Save the DBF path to settings
                    app_settings['quadra_last_dbf_path'] = dbf_path
                    window.metadata['_app_settings'] = app_settings
                    save_settings(app_settings)
                    
                    window["-STATUS_BAR-"].update(f"Załadowano plik DBF: {len(quadra_dbf_field_names)} pól wykrytych")
                except Exception as e:
                    window["-STATUS_BAR-"].update(f"Błąd odczytu pól DBF: {e}")
        
        elif event == "-QUADRA_CONFIG_MAPPING-":
            # Toggle visibility of mapping panel
            panel = window["-QUADRA_MAPPING_PANEL-"]
            # Ensure metadata is a dict (some GUI frameworks may initialize it as None)
            meta = getattr(panel, 'metadata', None)
            if meta is None or not isinstance(meta, dict):
                # initialize metadata to an empty dict to avoid AttributeError on .get
                panel.metadata = {}
                meta = panel.metadata
            current_visible = bool(meta.get('visible', False))
            new_visible = not current_visible
            panel.update(visible=new_visible)
            panel.metadata['visible'] = new_visible
        
        elif event == "-QUADRA_APPLY_MAPPING-":
            # Apply user-configured mapping and save to settings
            quadra_dbf_field_mapping = {}
            
            numer_field = values["-QUADRA_MAP_NUMER-"]
            if numer_field:
                quadra_dbf_field_mapping['numer_dbf'] = numer_field
            
            stawka_field = values["-QUADRA_MAP_STAWKA-"]
            if stawka_field:
                quadra_dbf_field_mapping['stawka'] = stawka_field
            
            czesci_field = values["-QUADRA_MAP_CZESCI-"]
            if czesci_field:
                quadra_dbf_field_mapping['czesci'] = czesci_field
            
            platnik_field = values["-QUADRA_MAP_PLATNIK-"]
            if platnik_field:
                quadra_dbf_field_mapping['platnik'] = platnik_field
            
            # Save mapping to settings
            app_settings = window.metadata.get('_app_settings', {})
            app_settings['quadra_dbf_field_mapping'] = quadra_dbf_field_mapping
            window.metadata['_app_settings'] = app_settings
            save_settings(app_settings)
            
            sg.popup(f"Mapowanie zastosowane:\n{quadra_dbf_field_mapping}", title="Mapowanie")
            window["-STATUS_BAR-"].update("Mapowanie pól DBF zastosowane i zapisane")
        
        elif event == "-QUADRA_RESET_MAPPING-":
            # Reset to auto-detection and clear saved mapping
            quadra_dbf_field_mapping = {}
            
            # Re-detect and set default values if DBF is loaded
            if quadra_dbf_field_names:
                numer_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_NUMER_FIELD_NAMES)
                stawka_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_STAWKA_FIELD_NAMES)
                czesci_field = detect_dbf_field_name(quadra_dbf_field_names, DBF_CZESCI_FIELD_NAMES)
                
                window["-QUADRA_MAP_NUMER-"].update(value=numer_field or '')
                window["-QUADRA_MAP_STAWKA-"].update(value=stawka_field or '')
                window["-QUADRA_MAP_CZESCI-"].update(value=czesci_field or '')
                window["-QUADRA_MAP_PLATNIK-"].update(value='')
            
            # Remove saved mapping from settings
            app_settings = window.metadata.get('_app_settings', {})
            if 'quadra_dbf_field_mapping' in app_settings:
                del app_settings['quadra_dbf_field_mapping']
            window.metadata['_app_settings'] = app_settings
            save_settings(app_settings)
            
            window["-STATUS_BAR-"].update("Mapowanie zresetowane do autodetekcji")

        elif event == "-QUADRA_CHECK_BTN-":
            # Validate inputs
            dbf_path = values["-QUADRA_DBF_PATH-"].strip()
            if not dbf_path:
                sg.popup_error("Wybierz plik DBF.")
                continue
            
            dbf_column = values["-QUADRA_DBF_COLUMN-"].strip()
            if not dbf_column:
                sg.popup_error("Podaj kolumnę DBF (np. B).")
                continue
            
            selected_spreadsheet = values["-QUADRA_SPREADSHEET_DROPDOWN-"]
            if not selected_spreadsheet:
                sg.popup_error("Wybierz arkusz kalkulacyjny z listy.")
                continue
            
            # Get spreadsheet ID
            try:
                combo_values = window["-QUADRA_SPREADSHEET_DROPDOWN-"].Values
                idx = combo_values.index(selected_spreadsheet)
                file_info = quadra_current_spreadsheets[idx]
                spreadsheet_id = file_info["id"]
                spreadsheet_name = file_info["name"]
            except (ValueError, IndexError):
                sg.popup_error("Błąd: nie można znaleźć wybranego arkusza.")
                continue
            
            # Get sheet names
            all_sheets = values["-QUADRA_ALL_SHEETS-"]
            if all_sheets:
                sheet_names = None  # Search all sheets
            else:
                selected_sheet = values["-QUADRA_SHEETS_DROPDOWN-"]
                if not selected_sheet:
                    sg.popup_error("Wybierz zakładkę lub zaznacz 'Wszystkie zakładki'.")
                    continue
                sheet_names = [selected_sheet]
            
            # Get search mode
            mode = 'substring' if values["-QUADRA_SUBSTRING-"] else 'exact'
            
            # Get column filter
            column_filter = values["-QUADRA_COLUMN_FILTER-"].strip()
            column_names = None
            if column_filter:
                # Split by comma/semicolon/newline
                column_names = [c.strip() for c in re.split(r'[,;\n]+', column_filter) if c.strip()]
            
            # Disable check button, enable stop
            window["-QUADRA_CHECK_BTN-"].update(disabled=True)
            window["-QUADRA_STOP_BTN-"].update(disabled=False)
            window["-STATUS_BAR-"].update(f"Sprawdzanie numerów z DBF w arkuszu {spreadsheet_name}...")
            
            # Start check thread
            global quadra_check_thread
            quadra_check_thread = threading.Thread(
                target=quadra_check_thread_func,
                args=(
                    window,
                    dbf_path,
                    dbf_column,
                    spreadsheet_id,
                    mode,
                    sheet_names,
                    column_names,
                    quadra_dbf_field_mapping if quadra_dbf_field_mapping else None
                ),
                daemon=True
            )
            quadra_check_thread.start()

        elif event == "-QUADRA_STOP_BTN-":
            quadra_stop_flag.set()
            window["-STATUS_BAR-"].update("Zatrzymywanie sprawdzania...")

        elif event == EVENT_QUADRA_CHECK_DONE:
            window["-QUADRA_CHECK_BTN-"].update(disabled=False)
            window["-QUADRA_STOP_BTN-"].update(disabled=True)
            
            results = values[EVENT_QUADRA_CHECK_DONE]
            if results == "error":
                window["-STATUS_BAR-"].update("Sprawdzanie zakończone z błędem.")
            else:
                # Display results in table
                table_data = []
                for result in results:
                    table_data.append(format_quadra_result_for_table(result))
                
                window["-QUADRA_RESULTS_TABLE-"].update(values=table_data)
                
                # Update status
                found_count = sum(1 for r in results if r['found'])
                missing_count = sum(1 for r in results if not r['found'])
                window["-QUADRA_STATUS-"].update(f"Znaleziono: {found_count} | Brakujących: {missing_count}")
                window["-STATUS_BAR-"].update(f"Sprawdzanie zakończone. Znaleziono: {found_count}, brakujących: {missing_count}")
                
                # Store results for export (preserve existing metadata)
                if not hasattr(window, 'metadata') or window.metadata is None:
                    window.metadata = {}
                window.metadata['quadra_results'] = results

        elif event == "-QUADRA_CLEAR_RESULTS-":
            window["-QUADRA_RESULTS_TABLE-"].update(values=[])
            window["-QUADRA_STATUS-"].update("Znaleziono: 0 | Brakujących: 0")
            window["-STATUS_BAR-"].update("Wyniki wyczyszczone.")
            if not hasattr(window, 'metadata') or window.metadata is None:
                window.metadata = {}
            window.metadata['quadra_results'] = []

        elif event == "-QUADRA_EXPORT_JSON-":
            results = window.metadata.get('quadra_results', []) if hasattr(window, 'metadata') else []
            if not results:
                sg.popup("Brak wyników do eksportu.", title="Eksport JSON")
                continue
            
            filename = sg.popup_get_file(
                "Zapisz wyniki do pliku JSON",
                save_as=True,
                default_extension=".json",
                file_types=(("JSON Files", "*.json"), ("All Files", "*.*")),
            )
            if filename:
                try:
                    export_data = export_quadra_results_to_json(results)
                    with open(filename, "w", encoding="utf-8") as f:
                        json.dump(export_data, f, ensure_ascii=False, indent=2)
                    sg.popup(f"Zapisano {len(results)} wyników do:\n{filename}", title="Eksport zakończony")
                    window["-STATUS_BAR-"].update(f"Wyniki zapisane do: {filename}")
                except Exception as e:
                    sg.popup_error(f"Błąd zapisu JSON: {e}")

        elif event == "-QUADRA_EXPORT_CSV-":
            results = window.metadata.get('quadra_results', []) if hasattr(window, 'metadata') else []
            if not results:
                sg.popup("Brak wyników do eksportu.", title="Eksport CSV")
                continue
            
            filename = sg.popup_get_file(
                "Zapisz wyniki do pliku CSV",
                save_as=True,
                default_extension=".csv",
                file_types=(("CSV Files", "*.csv"), ("All Files", "*.*")),
            )
            if filename:
                try:
                    csv_data = export_quadra_results_to_csv(results)
                    with open(filename, "w", encoding="utf-8") as f:
                        f.write(csv_data)
                    sg.popup(f"Zapisano {len(results)} wyników do:\n{filename}", title="Eksport zakończony")
                    window["-STATUS_BAR-"].update(f"Wyniki zapisane do: {filename}")
                except Exception as e:
                    sg.popup_error(f"Błąd zapisu CSV: {e}")

        # -------------------- Error handling --------------------
        elif event == EVENT_ERROR:
            error_msg = values[EVENT_ERROR]
            sg.popup_error(error_msg)
            window["-STATUS_BAR-"].update(f"Błąd: {error_msg}")

    # Save settings before closing
    if hasattr(window, 'metadata') and '_app_settings' in window.metadata:
        save_settings(window.metadata['_app_settings'])
    
    window.close()


if __name__ == "__main__":
    main()
