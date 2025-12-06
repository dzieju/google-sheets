# Google Sheets Search (Python)

Aplikacja CLI do przeszukiwania arkuszy Google należących do Twojego konta.

## Wymagania
- Python 3.9+
- Konto Google z włączonym API: Google Sheets API i Google Drive API
- Zainstalowane zależności (patrz requirements.txt)

## Przygotowanie
1. W Google Cloud Console:
   - Utwórz projekt, włącz Google Sheets API i Google Drive API.
   - Utwórz poświadczenia OAuth 2.0 typu "Desktop" (Installed app) i pobierz `credentials.json`.
2. Umieść `credentials.json` w katalogu projektu.
3. Zainstaluj zależności:
   ```
   python -m pip install -r requirements.txt
   ```
4. Pierwsze uruchomienie wywoła przeglądarkę (autoryzacja) i zapisze `token.json`.

## Użycie
- Listowanie arkuszy należących do Ciebie:
  ```
  python main.py list
  ```
- Przeszukiwanie (prostego substringu, case-insensitive):
  ```
  python main.py search "szukany tekst"
  ```
- Przeszukiwanie regex (przykład):
  ```
  python main.py search --regex --pattern "Faktura.*2025"
  ```

## Funkcjonalności wyszukiwania

### Wyszukiwanie w wielu kolumnach o tej samej nazwie
Aplikacja automatycznie wyszukuje **wszystkie kolumny** o podanej nazwie nagłówka w całym arkuszu:
- **W ramach jednego arkusza (zakładki)**: jeśli istnieje wiele kolumn o tej samej nazwie nagłówka (np. dwie kolumny "Numer zlecenia"), wszystkie będą przeszukane
- **Między różnymi zakładkami**: jeśli różne zakładki zawierają kolumny o tej samej nazwie, wszystkie będą uwzględnione
- **Kolejność wyników**: wyniki są zwracane w kolejności arkuszy (zakładek), a następnie w kolejności kolumn od lewej do prawej

### Dopasowanie nagłówków
- **Niewrażliwe na wielkość liter**: "Numer Zlecenia" = "numer zlecenia" = "NUMER ZLECENIA"
- **Ignorowanie białych znaków**: "Numer zlecenia" = "Numer  zlecenia" = " Numer zlecenia "
- **Normalizacja podkreślników**: "numer_zlecenia" = "numer zlecenia"

### Ignorowanie kolumn (pole "Ignoruj")
Aplikacja pozwala na wykluczenie określonych kolumn z wyszukiwania za pomocą pola "Ignoruj:":
- **Jak używać**: W interfejsie graficznym (GUI), w zakładce "Przeszukiwanie arkusza", znajduje się pole "Ignoruj" pod polem "Zapytanie"
- **Format**: Wprowadź nazwy kolumn do ignorowania, oddzielone przecinkami, średnikami lub nowymi liniami
- **Przykład**: `temp, debug, old` - ignoruje kolumny o nazwach "temp", "debug" i "old"
- **Wildcards**: Obsługuje proste wzorce z gwiazdką (*):
  - `temp*` - ignoruje kolumny zaczynające się na "temp" (np. "temporary", "temp_col")
  - `*old` - ignoruje kolumny kończące się na "old" (np. "test_old", "old")
  - `*debug*` - ignoruje kolumny zawierające "debug" (np. "debug_mode", "old_debug_temp")
- **Priorytet**: Kolumny pasujące do wzorców ignorowania nie będą zwracane nawet jeśli pasują do zapytania
- **Normalizacja**: Porównanie jest niewrażliwe na wielkość liter i białe znaki, tak jak w przypadku nazw kolumn

### Konfiguracja wierszy nagłówkowych (pole "Header rows")
Aplikacja pozwala na określenie, które wiersze arkusza zawierają nagłówki kolumn:
- **Jak używać**: W interfejsie graficznym (GUI), w zakładce "Przeszukiwanie arkusza", znajduje się pole "Header rows" (domyślnie "1")
- **Format**: Wprowadź numery wierszy oddzielone przecinkami (np. "1" dla pojedynczego wiersza, "1,2" dla dwóch wierszy)
- **Przykłady**:
  - `1` - tylko pierwszy wiersz zawiera nagłówki (domyślne)
  - `1,2` - nagłówki są w wierszach 1 i 2, wartości są łączone spacją dla każdej kolumny
  - `2` - tylko drugi wiersz zawiera nagłówki (dane zaczynają się od wiersza 3)
- **Łączenie nagłówków**: Gdy podano wiele wierszy (np. "1,2"), wartości z tych wierszy dla każdej kolumny są łączone spacją i normalizowane (trim + lowercase)
- **Przykład łączenia**: Jeśli wiersz 1 zawiera "First" a wiersz 2 zawiera "Last" w tej samej kolumnie, wynikowy nagłówek to "first last"

### Nazwa arkusza w wynikach
Wszystkie wyniki wyszukiwania i duplikatów zawierają teraz nazwę arkusza (zakładki):
- **W tabeli wyników**: Pierwsza kolumna "Arkusz" pokazuje nazwę zakładki (np. "2025")
- **W eksporcie JSON**: Pole "sheetName" zawiera nazwę zakładki dla każdego wyniku
- **Kompatybilność**: Zachowana pełna kompatybilność wsteczna - wszystkie wyniki zawierają nazwę arkusza

### Wykrywanie duplikatów
Funkcja wykrywania duplikatów również obsługuje wiele kolumn:
- Wykrywa duplikaty osobno w każdej kolumnie o podanej nazwie
- W wynikach rozróżnia kolumny dodając informację o literze kolumny (np. "Numer zlecenia (kolumna B)")

## Pliki
- main.py — CLI
- google_auth.py — obsługa OAuth i tworzenie klienta Sheets/Drive
- sheets_search.py — logika listowania i przeszukiwania arkuszy
- gui.py — graficzny interfejs użytkownika (GUI)

## Uwagi
- Aplikacja odczytuje całe arkusze (może być wolne przy bardzo dużych plikach). Możemy później dodać ograniczenia rozmiaru, progi lub przetwarzanie strumieniowe.
- Nie commituj `credentials.json` ani `token.json` do repo (dodane do .gitignore).

## Changelog

### v2.2 - Wiele wierszy nagłówkowych i nazwa arkusza w wynikach
- **NOWOŚĆ**: Dodano pole "Header rows" do konfiguracji wierszy nagłówkowych (domyślnie "1")
- **NOWOŚĆ**: Obsługa wielu wierszy nagłówkowych (np. "1,2") - wartości są łączone spacją dla każdej kolumny
- **NOWOŚĆ**: Kolumna "Arkusz" w tabelach wyników i duplikatów pokazuje nazwę zakładki
- **NOWOŚĆ**: Pole "sheetName" w eksporcie JSON zawiera nazwę zakładki dla każdego wyniku
- Zachowana pełna kompatybilność wsteczna - domyślne zachowanie (wiersz 1) pozostaje bez zmian

### v2.1 - Pole "Ignoruj" do wykluczania kolumn
- **NOWOŚĆ**: Dodano pole "Ignoruj:" w interfejsie GUI do wykluczania określonych kolumn z wyszukiwania
- **NOWOŚĆ**: Obsługa wielu wartości oddzielonych przecinkami, średnikami lub nowymi liniami
- **NOWOŚĆ**: Wsparcie dla wildcardów: `pattern*` (prefix), `*pattern` (suffix), `*pattern*` (contains)
- **NOWOŚĆ**: Kolumny pasujące do wzorców ignorowania nie są zwracane nawet jeśli pasują do zapytania
- Zachowana pełna kompatybilność wsteczna - puste pole "Ignoruj" zachowuje dotychczasowe zachowanie

### v2.0 - Obsługa wielu kolumn o tej samej nazwie
- **NOWOŚĆ**: Wyszukiwanie znajduje i przeszukuje wszystkie kolumny o podanej nazwie nagłówka, nie tylko pierwszą
- **NOWOŚĆ**: Obsługa kolumn o tej samej nazwie zarówno w ramach jednego arkusza jak i między zakładkami
- **NOWOŚĆ**: Wykrywanie duplikatów rozróżnia różne kolumny o tej samej nazwie
- Zachowana kompatybilność wsteczna - wszystkie wyniki są zwracane w spójnym formacie