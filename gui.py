"""
gui.py
GUI oparty na FreeSimpleGUI do przeszukiwania Google Sheets.
Wykorzystuje istniejące moduły google_auth i sheets_search.
"""

import json
import os
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
)


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
    """Format a single sheet search result as table row [Arkusz kalkulacyjny, Zakładka, Zlecenie, Stawka]."""
    return [
        result.get('spreadsheetName', ''),
        result.get('sheetName', ''),
        result.get('searchedValue', ''),
        result.get('stawka', ''),
    ]


def format_dup_result_for_table(result: dict) -> list:
    """Format a duplicate result as table row [Arkusz, Kolumna, Wartość, Ile razy, Przykładowe wiersze]."""
    rows = result.get('rows', [])
    # Show first 5 row numbers as example
    example_rows = ', '.join(str(r) for r in rows[:5])
    if len(rows) > 5:
        example_rows += f'... (+{len(rows) - 5})'
    
    return [
        f"{result.get('spreadsheetName', '')} / {result.get('sheetName', '')}",
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


def ss_search_thread_func(window, spreadsheet_id, spreadsheet_name, sheet_name, pattern, regex, case_sensitive, all_sheets=False, search_column_name=None, ignore_patterns=None):
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


def ss_search_all_spreadsheets_thread_func(window, pattern, regex, case_sensitive, search_column_name=None, ignore_patterns=None):
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
    # Definicja kolumn tabeli wyników
    table_headings = ["Arkusz kalkulacyjny", "Nazwa arkusza", "Zlecenie", "Stawka"]
    # Definicja kolumn tabeli duplikatów
    dup_table_headings = ["Arkusz / Zakładka", "Kolumna", "Wartość", "Ile razy", "Przykładowe wiersze"]
    
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
        [sg.Text("Ignoruj:", font=("Helvetica", 9, "bold"))],
        [sg.Text("Oddzielaj przecinkami, średnikami lub nową linią; wspiera '*' jako wildcard", font=("Helvetica", 8))],
        [sg.Multiline(key="-SHEET_IGNORE-", size=(None, 3), default_text="", expand_x=True)],
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


def create_layout():
    """Create main window layout with tabs."""
    tab_auth = sg.Tab("Autoryzacja", create_auth_tab())
    tab_files = sg.Tab("Pliki i podgląd", create_files_tab())
    tab_search = sg.Tab("Przeszukiwanie", create_search_tab())
    tab_single_sheet = sg.Tab("Przeszukiwanie arkusza", create_single_sheet_search_tab(), expand_x=True, expand_y=True)
    tab_settings = sg.Tab("Ustawienia", create_settings_tab())

    layout = [
        [sg.TabGroup([[tab_auth, tab_files, tab_search, tab_single_sheet, tab_settings]], key="-TABGROUP-", expand_x=True, expand_y=True)],
        [sg.StatusBar("Gotowe", key="-STATUS_BAR-", size=(60, 1))],
    ]
    return layout


# -------------------- Main GUI loop --------------------
def main():
    """Main function - runs the GUI event loop."""
    global drive_service, sheets_service, search_thread, ss_search_thread, dup_search_thread

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
                        ignore_patterns
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
                        ignore_patterns
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

        # -------------------- Error handling --------------------
        elif event == EVENT_ERROR:
            error_msg = values[EVENT_ERROR]
            sg.popup_error(error_msg)
            window["-STATUS_BAR-"].update(f"Błąd: {error_msg}")

    window.close()


if __name__ == "__main__":
    main()
