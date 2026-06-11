# EUPHIRE | Automatyczny Raport Sprzedaży "Labirynt Rozmów"

Ten projekt służy do automatycznego generowania pięknych i dynamicznych raportów sprzedażowych dla gry **Labirynt Rozmów** bezpośrednio z surowych eksportów transakcji w formacie `.xlsx`. Raporty są sformatowane według wytycznych **EUPHIRE Design System** i automatycznie publikowane online.

---

## 🔗 Link dla Marketera
Najnowsza wersja raportu jest zawsze dostępna pod stałym adresem:
👉 **[https://baciok91.github.io/euphire-raporty-sprzedazy/](https://baciok91.github.io/euphire-raporty-sprzedazy/)**

---

## 🚀 Instrukcja Obsługi Przyszłych Eksportów

Gdy pobierzesz nowy eksport zamówień i zechcesz zaktualizować raport:

1. **Wrzuć nowy plik `.xlsx`** bezpośrednio do tego folderu:
   `/Users/maciekckoklormam91/Desktop/Inne/Analiza sprzedazy EUPHIE/`

2. **Otwórz Terminal** w tym katalogu i uruchom poniższą komendę:
   ```bash
   ./.venv/bin/python generate_report.py
   ```
   *Skrypt automatycznie wykryje najnowszy plik zaczynający się na `Zamówienia` lub dowolny najnowszy arkusz `.xlsx`.*

3. **Gotowe!** Skrypt automatycznie wykona następujące kroki:
   - Wczyta i przeliczy najnowsze dane.
   - Zaktualizuje lokalne pliki `index.html` oraz `Raport_Labirynt_Euphire.html`.
   - Wyśle aktualizację na Twojego GitHuba, dzięki czemu **link dla marketera natychmiast zaktualizuje się do nowej wersji**.
   - Automatycznie otworzy podgląd raportu w Google Chrome na Twoim komputerze.

### 💡 Ręczne wskazanie pliku (Opcjonalnie)
Jeśli chcesz przeanalizować konkretny plik, a nie najnowszy, podaj jego nazwę jako argument:
```bash
./.venv/bin/python generate_report.py "Nazwa_Twojego_Pliku.xlsx"
```

---

## 🔒 Bezpieczeństwo i RODO/GDPR
Wszystkie pliki Excel (`*.xlsx`, `*.xls`) zawierające wrażliwe dane klientów (takie jak nazwiska, maile, adresy dostaw) zostały dodane do pliku `.gitignore`. 
**Surowe bazy danych nigdy nie zostaną wysłane na GitHub.** W repozytorium ląduje wyłącznie zanonimizowany, wygenerowany plik raportu HTML zawierający zagregowane wykresy i statystyki.
