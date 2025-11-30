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
from sheets_search import list_spreadsheets_owned_by_me, search_in_spreadsheets


# -------------------- Events --------------------
EVENT_AUTH_DONE = "-AUTH_DONE-"
EVENT_FILES_LOADED = "-FILES_LOADED-"
EVENT_SHEETS_LOADED = "-SHEETS_LOADED-"
EVENT_PREVIEW_LOADED = "-PREVIEW_LOADED-"
EVENT_SEARCH_RESULT = "-SEARCH_RESULT-"
EVENT_SEARCH_DONE = "-SEARCH_DONE-"
EVENT_ERROR = "-ERROR-"

# -------------------- Global state --------------------
drive_service = None
sheets_service = None
current_spreadsheets = []
search_thread = None
stop_search_flag = threading.Event()


# -------------------- Helper functions --------------------
def format_result(result: dict) -> str:
    """Format a single search result for display."""
    return (
        f"[{result['spreadsheetName']}] "
        f"Arkusz: {result['sheetName']}, "
        f"Komórka: {result['cell']}, "
        f"Wartość: {result['value']}"
    )


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
            max_files=max_files if max_files and max_files > 0 else None,
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


def create_layout():
    """Create main window layout with tabs."""
    tab_auth = sg.Tab("Autoryzacja", create_auth_tab())
    tab_files = sg.Tab("Pliki i podgląd", create_files_tab())
    tab_search = sg.Tab("Przeszukiwanie", create_search_tab())
    tab_settings = sg.Tab("Ustawienia", create_settings_tab())

    layout = [
        [sg.TabGroup([[tab_auth, tab_files, tab_search, tab_settings]], key="-TABGROUP-")],
        [sg.StatusBar("Gotowe", key="-STATUS_BAR-", size=(60, 1))],
    ]
    return layout


# -------------------- Main GUI loop --------------------
def main():
    """Main function - runs the GUI event loop."""
    global drive_service, sheets_service, search_thread

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
                # Parse selected item to get ID
                selected_text = selected[0]
                # Extract ID from "name  (id)" format
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
            # Append to multiline
            current_text = window["-SEARCH_RESULTS-"].get()
            new_line = format_result(result)
            if current_text:
                window["-SEARCH_RESULTS-"].update(current_text + "\n" + new_line)
            else:
                window["-SEARCH_RESULTS-"].update(new_line)
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

        # -------------------- Error handling --------------------
        elif event == EVENT_ERROR:
            error_msg = values[EVENT_ERROR]
            sg.popup_error(error_msg)
            window["-STATUS_BAR-"].update(f"Błąd: {error_msg}")

    window.close()


if __name__ == "__main__":
    main()
