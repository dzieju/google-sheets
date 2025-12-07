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

### Ignorowanie kolumn i wartości (pole "Ignoruj")
Aplikacja pozwala na wykluczenie określonych kolumn oraz wartości komórek z wyników wyszukiwania za pomocą pola "Ignoruj:":
- **Jak używać**: W interfejsie graficznym (GUI), w zakładce "Przeszukiwanie arkusza", znajduje się pole "Ignoruj" pod polem "Zapytanie"
- **Format**: Wprowadź wzorce do ignorowania, oddzielone przecinkami, średnikami lub nowymi liniami
- **Przykład**: `temp, debug, old` - ignoruje kolumny o nazwach zawierających "temp", "debug" lub "old", oraz wartości komórek zawierające te słowa
- **Zastosowanie**:
  - **Ignorowanie nagłówków kolumn**: Kolumny których nazwy pasują do wzorców nie będą przeszukiwane
  - **Ignorowanie wartości komórek**: Dopasowane wartości komórek zawierające wzorce ignorowania będą pomijane w wynikach
- **Wzorce dopasowania**:
  - Wzorce **bez gwiazdki** (`*`) są dopasowywane jako **podciąg** (substring, case-insensitive):
    - `https` - ignoruje wszystkie wartości zawierające "https" (np. URL-e "https://...")
    - `temp` - ignoruje kolumny i wartości zawierające "temp" (np. "temporary", "temp_data", "123temp")
  - Wzorce **z gwiazdką** (`*`) działają jako **wildcard**:
    - `temp*` - ignoruje kolumny/wartości zaczynające się na "temp" (np. "temporary", "temp_col")
    - `*old` - ignoruje kolumny/wartości kończące się na "old" (np. "test_old", "old")
    - `*debug*` - ignoruje kolumny/wartości zawierające "debug" (np. "debug_mode", "old_debug_temp")
- **Priorytet**: Kolumny i wartości pasujące do wzorców ignorowania nie będą zwracane nawet jeśli pasują do zapytania
- **Normalizacja**: Porównanie jest niewrażliwe na wielkość liter i białe znaki
- **Przykład użycia**: Wpisanie "https" w polu Ignoruj spowoduje pominięcie wszystkich wyników zawierających URL-e z "https://"

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

## Funkcja Quadra - Sprawdzanie numerów zleceń z DBF

Zakładka "Quadra" pozwala na porównanie numerów zleceń z pliku DBF z zawartością arkuszy Google Sheets oraz wyświetlanie dodatkowych informacji z DBF (stawka, części).

### Jak używać Quadra

1. **Wybierz plik DBF**:
   - Kliknij "Wybierz plik DBF" aby wskazać plik z numerami zleceń
   - Domyślnie aplikacja wczytuje kolumnę B (można zmienić w polu "Kolumna DBF")
   - Aplikacja automatycznie wykrywa i wczytuje dodatkowe pola: stawka i części

2. **Podaj kolumnę DBF**:
   - Wpisz literę kolumny (np. "B", "C") lub numer (np. "1", "2")
   - Domyślnie: "B" (druga kolumna)

3. **Wybierz arkusz kalkulacyjny**:
   - Kliknij "Odśwież listę arkuszy" aby załadować dostępne arkusze
   - Wybierz arkusz z listy rozwijanej

4. **Wybierz zakładki do przeszukania**:
   - Zaznacz "Wszystkie zakładki" aby przeszukać wszystkie (domyślnie)
   - Lub wybierz konkretną zakładkę z listy

5. **Opcje wyszukiwania**:
   - **Exact**: Dopasowanie dokładne (ignoruje białe znaki, wielkość liter, porównuje liczby)
   - **Substring**: Dopasowanie podciągu (jedna wartość zawiera drugą)
   - **Kolumny do przeszukania**: Pozostaw puste aby przeszukać wszystkie kolumny, lub podaj nazwy kolumn oddzielone przecinkami

6. **Sprawdź**:
   - Kliknij "Sprawdź" aby rozpocząć porównanie
   - Wyniki pojawią się w tabeli z następującymi kolumnami (w kolejności):
     - **Arkusz**: Nazwa zakładki gdzie znaleziono wartość
     - **Numer z DBF**: Wartość z pliku DBF (główna kolumna wyszukiwania)
     - **Stawka**: Wartość stawki z DBF
     - **Czesci**: Wartość części z DBF
     - **Status**: "Found" (znaleziono) lub "Missing" (brakuje)
     - **Kolumna**: Nazwa kolumny gdzie znaleziono wartość
     - **Wiersz**: Numer wiersza gdzie znaleziono wartość
     - **Uwagi**: Dodatkowe informacje

7. **Eksport wyników**:
   - **Eksportuj JSON**: Zapisz wyniki w formacie JSON (zawiera wszystkie pola: numer_dbf, stawka, czesci)
   - **Eksportuj CSV**: Zapisz wyniki w formacie CSV (zawiera wszystkie pola: numer_dbf, stawka, czesci)

### Konfiguracja mapowania pól DBF

Po wybraniu pliku DBF możesz skonfigurować mapowanie pól:

1. **Automatyczne wykrywanie** (domyślne):
   - **Numer z DBF**: Wykrywany z pól: `NUMER`, `NUMBER`, `NR`, `ORDER`, `ZLECENIE`
   - **Stawka**: Wykrywana z pól: `STAWKA`, `STAW`, `RATE`, `PRICE`, `CENA`
   - **Części**: Wykrywana z pól: `CZESCI`, `PARTS`, `CZESC`, `PART`
   
2. **Ręczna konfiguracja**:
   - Kliknij "Konfiguruj mapowanie pól" aby otworzyć panel konfiguracji
   - Wybierz pola DBF z list rozwijanych dla każdego typu danych
   - Kliknij "Zastosuj mapowanie" aby użyć własnego mapowania
   - Kliknij "Resetuj" aby wrócić do autodetekcji

### Wykrywanie dodatkowych pól z DBF

Aplikacja automatycznie wykrywa i wczytuje dodatkowe pola z pliku DBF:

- **Numer z DBF**: Wykrywany z pól o nazwach: `NUMER`, `NUMBER`, `NR`, `ORDER`, `ZLECENIE` (case-insensitive)
  - Definiowane w stałej: `DBF_NUMER_FIELD_NAMES` w `quadra_service.py`
- **Stawka**: Wykrywana z pól o nazwach: `STAWKA`, `STAW`, `RATE`, `PRICE`, `CENA` (case-insensitive)
  - Definiowane w stałej: `DBF_STAWKA_FIELD_NAMES` w `quadra_service.py`
- **Części**: Wykrywana z pól o nazwach: `CZESCI`, `PARTS`, `CZESC`, `PART` (case-insensitive)
  - Definiowane w stałej: `DBF_CZESCI_FIELD_NAMES` w `quadra_service.py`

Jeśli pola nie istnieją w DBF, wyświetlane są puste wartości (bez błędu).

**Uwaga dla deweloperów**: Aby dodać nowe alternatywne nazwy pól, edytuj stałe `DBF_NUMER_FIELD_NAMES`, `DBF_STAWKA_FIELD_NAMES` i `DBF_CZESCI_FIELD_NAMES` w pliku `quadra_service.py`.

### Przykładowe użycie

```
1. Masz plik DBF z kolumnami:
   - ORDER (kolumna B): 12345, 67890, ABC-001
   - STAWKA: 150.00, 200.50, 100.00
   - CZESCI: ABC, XYZ, DEF
2. Chcesz sprawdzić czy te numery są w arkuszu Google "Zlecenia 2025"
3. W zakładce Quadra:
   - Wybierz plik DBF
   - Zostaw "B" jako kolumnę
   - Wybierz arkusz "Zlecenia 2025"
   - Zaznacz "Wszystkie zakładki"
   - Kliknij "Sprawdź"
4. Wyniki pokażą:
   - Które numery zostały znalezione i gdzie
   - Stawkę z DBF dla każdego numeru
   - Części z DBF dla każdego numeru
   - Które numery brakują
```

### Dopasowanie wartości

- **Tryb Exact** (domyślny):
  - Ignoruje białe znaki i wielkość liter
  - Porównuje wartości numeryczne jako liczby (np. "12345" = 12345)
  - Normalizuje formatowanie liczb (usuwa spacje, separatory tysięcy)

- **Tryb Substring**:
  - Sprawdza czy jedna wartość zawiera drugą
  - Przydatne gdy numery w arkuszu mają dodatkowe przedrostki/przyrostki

### Zapisywanie wyników do Google Sheets

Możesz zapisać wyniki sprawdzenia do arkusza Google Sheets używając funkcji `write_quadra_results_to_sheet`:

- Dane zapisywane są do kolumn I (Stawka) i J (Czesci)
- Kolumny I i J są automatycznie tworzone/aktualizowane
- Pierwszy wiersz zawiera nagłówki: "Stawka" i "Czesci"
- Zachowana jest kompatybilność z istniejącymi danymi w innych kolumnach

### Struktura eksportowanych danych

**JSON**:
```json
[
  {
    "dbfValue": "12345",
    "stawka": "150.00",
    "status": "Found",
    "sheetName": "2025",
    "columnName": "Numer zlecenia",
    "columnIndex": 1,
    "rowIndex": 5,
    "matchedValue": "12345",
    "czesci": "ABC",
    "notes": "Found in 2025 at B6"
  },
  {
    "dbfValue": "99999",
    "stawka": "200.50",
    "status": "Missing",
    "sheetName": "",
    "columnName": "",
    "columnIndex": null,
    "rowIndex": null,
    "matchedValue": "",
    "czesci": "XYZ",
    "notes": "Missing"
  }
]
```

**CSV**:
```
DBF_Value,Stawka,Status,SheetName,ColumnName,ColumnIndex,RowIndex,MatchedValue,Czesci,Notes
12345,150.00,Found,2025,Numer zlecenia,1,5,12345,ABC,Found in 2025 at B6
99999,200.50,Missing,,,,,XYZ,Missing
```

## Pliki
- main.py — CLI
- google_auth.py — obsługa OAuth i tworzenie klienta Sheets/Drive
- sheets_search.py — logika listowania i przeszukiwania arkuszy
- gui.py — graficzny interfejs użytkownika (GUI)
- quadra_service.py — moduł do sprawdzania numerów z DBF w arkuszach Google Sheets

## Uwagi
- Aplikacja odczytuje całe arkusze (może być wolne przy bardzo dużych plikach). Możemy później dodać ograniczenia rozmiaru, progi lub przetwarzanie strumieniowe.
- Nie commituj `credentials.json` ani `token.json` do repo (dodane do .gitignore).

## Changelog

### v3.2 - Quadra: Konfiguracja mapowania pól DBF i zmiana kolejności kolumn
- **NOWOŚĆ**: Panel konfiguracji mapowania pól DBF
  - Przycisk "Konfiguruj mapowanie pól" otwiera panel z listami rozwijanymi
  - Możliwość ręcznego wyboru pól DBF dla: Numer z DBF, Stawka, Czesci
  - Przyciski "Zastosuj mapowanie" i "Resetuj" do zarządzania konfiguracją
  - Automatyczne załadowanie listy pól po wybraniu pliku DBF
- **NOWOŚĆ**: Rozszerzone automatyczne wykrywanie pól DBF
  - Numer z DBF: NUMER, NUMBER, NR, ORDER, ZLECENIE
  - Stawka: STAWKA, STAW, RATE, PRICE, CENA (dodano CENA)
  - Części: CZESCI, PARTS, CZESC, PART (dodano CZESC, PART)
- **ZMIANA**: Nowa kolejność kolumn w tabeli wyników (Czesci obok Stawka)
  - Arkusz, Numer z DBF, Stawka, Czesci, Status, Kolumna, Wiersz, Uwagi
  - Wcześniej: Numer z DBF, Stawka, Status, Arkusz, Kolumna, Wiersz, Czesci_extra, Uwagi
- **ZMIANA**: Zmiana nazwy kolumny z "Czesci_extra" na "Czesci"
  - Zmiana w GUI, eksporcie JSON/CSV i zapisie do Google Sheets
  - Kolumna J w Google Sheets ma teraz nagłówek "Czesci" zamiast "Czesci_extra"
- **NOWOŚĆ**: Funkcja `get_dbf_field_names()` - odczyt nazw pól z DBF
- **NOWOŚĆ**: Funkcja `map_dbf_record_to_result()` zwraca teraz również pole 'numer_dbf'
  - Obsługa opcjonalnego parametru `mapping` dla ręcznej konfiguracji
  - Kompatybilność wsteczna z istniejącym kodem
- **NOWOŚĆ**: Rozbudowane testy jednostkowe (47 testów, wszystkie przechodzą)
  - Testy mapowania użytkownika
  - Testy alternatywnych nazw pól (CENA, CZESC, PART)
  - Testy nowej kolejności kolumn w tabeli
- Zachowana pełna kompatybilność wsteczna

### v3.1 - Quadra: Dodatkowe kolumny Stawka i Czesci z DBF
- **NOWOŚĆ**: Automatyczne wykrywanie i wyświetlanie kolumny "Stawka" z DBF
  - Wykrywanie z alternatywnych nazw pól: STAWKA, STAW, RATE, PRICE (case-insensitive)
  - Kolumna "Stawka" pojawia się w GUI zaraz po "Numer z DBF"
  - Wartości zapisywane do kolumny I (index 8) w Google Sheets
- **NOWOŚĆ**: Automatyczne wykrywanie i wyświetlanie kolumny "Czesci_extra" z DBF
  - Wykrywanie z alternatywnych nazw pól: CZESCI, PARTS (case-insensitive)
  - Kolumna "Czesci_extra" pojawia się w GUI obok kolumny "Czesci"
  - Wartości zapisywane do kolumny J (index 9) w Google Sheets
- **NOWOŚĆ**: Funkcja `write_quadra_results_to_sheet()` do zapisu wyników do Google Sheets
  - Automatyczne tworzenie/aktualizacja kolumn I i J
  - Nagłówki "Stawka" i "Czesci_extra" w pierwszym wierszu
  - Zachowana kompatybilność z istniejącymi danymi
- **NOWOŚĆ**: Nowe funkcje pomocnicze w `quadra_service.py`:
  - `detect_dbf_field_name()` - wykrywanie nazw pól z alternatyw
  - `map_dbf_record_to_result()` - mapowanie rekordów DBF na wyniki
  - `read_dbf_records_with_extra_fields()` - odczyt DBF z dodatkowymi polami
- **NOWOŚĆ**: Eksport JSON i CSV zawiera teraz pola "stawka" i "czesci"
- **NOWOŚĆ**: Rozbudowane testy jednostkowe (43 testy, wszystkie przechodzą)
  - Testy wykrywania pól z alternatywnych nazw
  - Testy obsługi brakujących pól
  - Testy formatowania i eksportu z nowymi kolumnami
  - Testy zapisu do Google Sheets (z mockami)
- Bezpieczna obsługa brakujących pól - puste stringi zamiast błędów
- Pełna kompatybilność wsteczna - obsługa zarówno prostych wartości jak i słowników z rekordami

### v3.0 - Quadra: Sprawdzanie numerów zleceń z DBF
- **NOWOŚĆ**: Dodano zakładkę "Quadra" do sprawdzania numerów zleceń z plików DBF
- **NOWOŚĆ**: Obsługa plików DBF - wczytywanie wartości z wybranej kolumny (domyślnie B)
- **NOWOŚĆ**: Porównywanie wartości z DBF z zawartością arkuszy Google Sheets
- **NOWOŚĆ**: Dwa tryby dopasowania: Exact (dokładne, z normalizacją) i Substring (podciąg)
- **NOWOŚĆ**: Wyniki pokazują status (Found/Missing), arkusz, kolumnę, wiersz i uwagi
- **NOWOŚĆ**: Eksport wyników do JSON i CSV
- **NOWOŚĆ**: Możliwość ograniczenia wyszukiwania do wybranych zakładek i kolumn
- **NOWOŚĆ**: Moduł `quadra_service.py` z logiką obsługi DBF i porównywania
- Dodano 23 testy jednostkowe dla modułu Quadra
- Dodano zależność `dbfread>=2.0.7` do requirements.txt

### v2.3 - Ignorowanie wartości komórek oraz nowa semantyka wzorców
- **NOWOŚĆ**: Pole "Ignoruj" filtruje teraz również **wartości dopasowanych komórek**, nie tylko nagłówki kolumn
- **ZMIANA**: Wzorce bez gwiazdki (`*`) są teraz dopasowywane jako **podciąg** (substring match), nie jako exact match
  - Przykład: `https` w polu Ignoruj pominie wszystkie wyniki zawierające "https" w wartości (np. URL-e)
  - Przykład: `temp` pominie zarówno kolumny jak i wartości zawierające "temp"
- **ZACHOWANE**: Wzorce z gwiazdką (`*`) działają jak wcześniej (prefix/suffix/contains wildcard)
- **NOWOŚĆ**: Nowa funkcja `matches_ignore_value()` do sprawdzania wartości komórek
- Dodano kompleksowe testy jednostkowe dla nowej funkcjonalności
- Zachowana pełna kompatybilność wsteczna dla wildcard patterns

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