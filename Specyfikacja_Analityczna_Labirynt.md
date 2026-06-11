# Specyfikacja Analityczna: Sprzedaż "Labirynt Rozmów" (EUPHIRE)

Ten plik stanowi stałą i powtarzalną specyfikację analityczną (wytyczne) dla wyciągania i prezentowania statystyk sprzedażowych, dedykowanych wyłączenie grze **Labirynt Rozmów**. Został sformatowany tak, aby z łatwością instruować sztuczną inteligencję (agenta roboczego) podczas wczytywania nowych zrzutów z bazy zamówień.

## 1. Cel Raportowania
Generowanie w pełni zautomatyzowanego wizualnego raportu dynamicznego `.html` (z użyciem bibliotek typu Chart.js), ukazującego trendy dzienne i tygodniowe, skoncentrowanego na wolumenie sprzedażowym kluczowego produktu marki.

## 2. Plik Zródłowy (Do aktualizacji za każdym razem)
Wskaż agentowi surowy zrzut transakcji, np.:
`Plik: /Users/maciekckoklormam91/Desktop/Analiza sprzedazy EUPHIE/Zamówienia_NOWY_MIESIAC.xlsx`

## 3. Logika Filtrowania i Mapowania (Wymagane!)
System musi kategorycznie wykluczyć całą resztę sprzedaży (np. webinary, warsztaty) i zawęzić zestawienie wyłącznie do poniższych, precyzyjnie brzmiących wierszy w kolumnie **Nazwa Kursu**:
1. `Labirynt Rozmów | Gra Psychologiczna dla Par i Przyjaciół [Gra karciana]` – mapuj na tzw. **Wersję Fizyczną (Pudełko)**
2. `Labirynt Rozmów [GRA ON-LINE]` – mapuj na **Wersję Cyfrową (On-line)**
3. `Labirynt Rozmów [Pudełko Premium + GRA ON-LINE]` – mapuj na **Pakiet (Pudełko + On-line)**

*Zadaniem skryptu jest oczyszczenie uszkodzonych wierszy z Excela (typu "Suma:" na końcu formularza) przez odrzucanie rekordów bez daty (lub z niepoprawnym `NaT`).*

## 4. Analiza Metryk
Program analityczny z każdorazowym odpaleniem musi wygenerować:
- **Metryki Całkowite:** Sumaryczny Przychód Brutto (tylko z tych wariantów), Łączna Sprzedana Ilość podziałem na pakiety/wydania, Ilość Zamówień, AOV dla wariantów Labiryntu.
- **Wykresy Dynamiki:**
  - Liniowy wykres dzienny: dzienne szczyty sprzedażowe dla połączonych wyników wszystkich trzech pakietów.
  - Słupkowy wykres tygodniowy: ujęcie długoterminowe.
  - Opcjonalnie: Kołowy wykres rozkładu procentowego (który wariant generuje największy przypływ gotówki).

## 5. Wytyczne Wizualne i Brand Identity: EUPHIRE Design System
Agent kodujący widok ma **bezwzględny zakaz** używania domyślnych, przypadkowych kolorów. Trzymaj się ram marki EUPHIRE:
- **Układ Typefaces (Z Google Fonts):**
  - Główne Headery (H1-H3): `font-family: 'Montserrat', sans-serif; font-weight: 600;`
  - Paragrafy, analiza opisowa: `font-family: 'Merriweather', serif; font-weight: 400;` (Ważne fragmenty kursywą `italic`).
  - Drobne wpisy, daty, liczby, informacje ze stopki: `font-family: 'Roboto Mono', monospace; font-weight: 400;`
- **Kolorystyka Bazowa:**
  - Tło główne: **Frost White** (`#FAFAFA`)
  - Strefa Nagłówka (Pasek tytułowy) / Tła dla kontrastowych metryk: **Evergreen** (`#004D54`)
  - Akcenty ostrzegawcze i mocne punkty wizualne: **Ember** (`#FCAE2F`)
  - Czcionki bazowe na jasnym tle: **Obsidian Black** (`#1F1F1F`)
  - Tła pomocnicze elementów, ramki pudełek i subtelności: **Mist** (`#B2CACC`) i **Nocturne** (`#002E32`) na potrzeby gradientów.
- **Zasady Logo:** W lewym / centralnym rogu headera osadź podstawową wersję logo SVG o adresie bezpośrednim: `https://euphire.pl/wp-content/uploads/2025/11/logo_primary.svg`

## Instrukcja do Wklejenia Modelowi Językowemu (Prompt):
*"Użyj pliku `Specyfikacja_Analityczna_Labirynt.md` żeby nadpisać skrypt raportowania pod nowy plik Excel: `<podaj-tu-adres-xlsx>`. Zastosuj EUPHIRE Design System zdefiniowany w tym szablonie, przelicz ponownie i odpal automatycznie docelowy `Raport_Labirynt_Euphire.html` w Chrome."*
