# Google Sheets Search (Python)

Aplikacja CLI do przeszukiwania arkuszy Google należących do Twojego konta.

Wymagania
- Python 3.9+
- Konto Google z włączonym API: Google Sheets API i Google Drive API
- Zainstalowane zależności (patrz requirements.txt)

Przygotowanie
1. W Google Cloud Console:
   - Utwórz projekt, włącz Google Sheets API i Google Drive API.
   - Utwórz poświadczenia OAuth 2.0 typu "Desktop" (Installed app) i pobierz `credentials.json`.
2. Umieść `credentials.json` w katalogu projektu.
3. Zainstaluj zależności:
   ```
   python -m pip install -r requirements.txt
   ```
4. Pierwsze uruchomienie wywoła przeglądarkę (autoryzacja) i zapisze `token.json`.

Użycie
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

Pliki
- main.py — CLI
- google_auth.py — obsługa OAuth i tworzenie klienta Sheets/Drive
- sheets_search.py — logika listowania i przeszukiwania arkuszy

Uwagi
- Aplikacja odczytuje całe arkusze (może być wolne przy bardzo dużych plikach). Możemy później dodać ograniczenia rozmiaru, progi lub przetwarzanie strumieniowe.
- Nie commituj `credentials.json` ani `token.json` do repo (dodane do .gitignore).