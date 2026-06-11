# Instrukcja Obsługi i Przepływu (Workflow) dla Agenta AI

Ten plik instruuje dowolnego przyszłego Agenta AI (np. Antigravity), jak automatycznie obsłużyć proces aktualizacji raportu po wrzuceniu przez użytkownika nowego pliku eksportu zamówień z WooCommerce.

---

## 🎯 Cel Workflow
Zautomatyzowana aktualizacja interaktywnego panelu sprzedażowego dla gry **Labirynt Rozmów** (EUPHIRE) i natychmiastowa publikacja na GitHub Pages.

---

## 💻 Instrukcja Krok po Kroku (Co musi zrobić Agent)

Gdy użytkownik poinformuje Cię, że dodał nowy eksport danych (plik `.xlsx`) lub poprosi o aktualizację raportu:

### Krok 1: Zlokalizuj najnowszy plik Excel
Nowy plik powinien znajdować się w głównym katalogu roboczym:
`/Users/maciekckoklormam91/Desktop/Inne/Analiza sprzedazy EUPHIE`
Skrypt `generate_report.py` automatycznie wykryje najnowszy plik (posortowany po czasie modyfikacji), który zawiera słowo `am*wienia` lub rozszerzenie `.xlsx` (z wyłączeniem plików tymczasowych).

### Krok 2: Uruchom skrypt generujący raport
Wykonaj komendę w terminalu:
```bash
./.venv/bin/python generate_report.py
```
*Uwaga: Nie ma potrzeby ręcznego otwierania przeglądarki Chrome na maszynie użytkownika. Skrypt wykona całą pracę w tle i wypchnie kod na produkcję.*

### Krok 3: Wyświetl link LIVE użytkownikowi
Po pomyślnym uruchomieniu skryptu, na samym końcu swojej odpowiedzi **obowiązkowo wypisz link do raportu online**:
👉 **[https://baciok91.github.io/euphire-raporty-sprzedazy/](https://baciok91.github.io/euphire-raporty-sprzedazy/)**

---

## ⚙️ Jak Działa Skrypt `generate_report.py`?
Skrypt automatycznie wykonuje następujące operacje:
1. **Wczytanie danych**: Ładuje plik `.xlsx` od 3. wiersza (nagłówki są w wierszu 3, `header=2` w pandas).
2. **Czyszczenie**: Odrzuca wiersze z pustymi/uszkodzonymi datami (np. podsumowania tabeli).
3. **Filtrowanie**: Zawęża zbiór wyłącznie do gry "Labirynt Rozmów" w wariantach:
   - `Wersja Fizyczna (Pudełko)` (Nazwa kursu: `Labirynt Rozmów | Gra Psychologiczna dla Par i Przyjaciół [Gra karciana]`)
   - `Wersja Cyfrową (On-line)` (Nazwa kursu: `Labirynt Rozmów [GRA ON-LINE]`)
   - `Pakiet (Pudełko + On-line)` (Nazwa kursu: `Labirynt Rozmów [Pudełko Premium + GRA ON-LINE]`)
4. **Wstrzykiwanie danych**: Konwertuje dane na zanonimizowaną listę JSON i wstrzykuje ją jako `rawTransactions` do kodu JS w pliku HTML.
5. **Kompilacja**: Nadpisuje pliki `index.html` oraz `Raport_Labirynt_Euphire.html`.
6. **Git / Deploy**: Wykonuje `git add`, `git commit` z opisem pliku wejściowego, a następnie `git push origin main` do dedykowanego repozytorium GitHub Pages.

---

## 🛠 Rozwiązywanie Problemów i Wytyczne dla Agenta (Troubleshooting)

### 1. Błąd braku bibliotek Python
Jeśli skrypt nie uruchamia się z powodu brakujących paczek (np. `pandas`, `openpyxl`), zainstaluj je w środowisku wirtualnym:
```bash
./.venv/bin/pip install pandas openpyxl
```

### 2. Zmiana struktury kolumn w WooCommerce
Skrypt oczekuje w pliku Excel dokładnie takich kolumn (wymagany nagłówek w wierszu index 2):
`['Numer', 'ID Zamówienia', 'Data', 'Dane Klienta', 'Dane do faktury', 'Nazwa Kursu', 'Sztuk', 'Kwota Brutto', 'Kwota Netto', 'Podatek', '% VAT', 'Typ']`
Jeśli kolumny się zmienią, dopasuj mapowanie w skrypcie `generate_report.py` tak, aby poprawnie pobierał:
- `Data` (do filtrowania dat i wykresów)
- `Nazwa Kursu` (do rozróżniania wariantów)
- `Kwota Brutto` (jest to rzeczywista kwota wpłacona przez klienta z uwzględnieniem rabatów)
- `Sztuk` (wolumen sprzedaży)

### 3. Błąd uwierzytelnienia Git (Push Error)
Jeśli `git push` zakończy się błędem uwierzytelnienia:
- Upewnij się, czy token lub sesja użytkownika w CLI GitHub (`gh`) nie wygasła.
- Jeśli jesteś w stanie, spróbuj użyć CLI `gh auth status` do sprawdzenia stanu zalogowania.
- W ostateczności poinformuj użytkownika o błędzie i poproś o lokalny restart uwierzytelnienia Git.
