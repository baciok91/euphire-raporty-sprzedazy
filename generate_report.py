import pandas as pd
import json
import os
import sys
import glob
from datetime import datetime
import subprocess

# Auto-detekcja pliku bazy danych lub pobranie z argumentu
script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
os.chdir(script_dir)

if len(sys.argv) > 1:
    file_path = sys.argv[1]
else:
    # Szukaj plików o nazwie pasującej do "Zamówienia" lub "Zamówienia"
    excel_files = glob.glob('*am*wienia*.xlsx')
    if not excel_files:
        # Alternatywnie dowolny .xlsx oprócz starego pliku bazowego i plików tymczasowych
        excel_files = [f for f in glob.glob('*.xlsx') if f != 'od wrzesnia do maja.xlsx' and not f.startswith('~$')]
    
    if excel_files:
        # Posortuj po czasie modyfikacji (najnowszy na początku)
        excel_files.sort(key=os.path.getmtime, reverse=True)
        file_path = excel_files[0]
    else:
        file_path = 'od wrzesnia do maja.xlsx'

print(f"Używam pliku bazy danych: {file_path}")

# Wczytaj dane
df = pd.read_excel(file_path, header=2)

# Oczyść wiersze z uszkodzoną datą
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
df = df.dropna(subset=['Data']).copy()
df['Kwota Brutto'] = pd.to_numeric(df['Kwota Brutto'], errors='coerce').fillna(0)
df['Sztuk'] = pd.to_numeric(df['Sztuk'], errors='coerce').fillna(0)

# -----------------
# FILTROWANIE TYLKO DLA "Labirynt Rozmów"
# -----------------
TARGET_COURSES = {
    'Labirynt Rozmów | Gra Psychologiczna dla Par i Przyjaciół [Gra karciana]': 'Wersja Fizyczna (Pudełko)',
    'Labirynt Rozmów [GRA ON-LINE]': 'Wersja Cyfrową (On-line)',
    'Labirynt Rozmów [Pudełko Premium + GRA ON-LINE]': 'Pakiet (Pudełko + On-line)'
}

df_filtered = df[df['Nazwa Kursu'].isin(TARGET_COURSES.keys())].copy()
df_filtered['Wariant'] = df_filtered['Nazwa Kursu'].map(TARGET_COURSES)

# Statystyki ogólne dla Labiryntu (wypis w konsoli)
total_revenue = df_filtered['Kwota Brutto'].sum()
total_sales = df_filtered['Sztuk'].sum()
total_orders = len(df_filtered)
print(f"Dane wczytane pomyślnie. Suma przychodów: {total_revenue:.2f} zł, Liczba zamówień: {total_orders}, Sztuk: {total_sales}")

# Przygotuj surowe dane transakcji do wstrzyknięcia do JS (zanonimizowane)
df_filtered['Dzien'] = df_filtered['Data'].dt.strftime('%Y-%m-%d')
raw_transactions = df_filtered[['Dzien', 'Wariant', 'Kwota Brutto', 'Sztuk']].copy()
raw_transactions_list = raw_transactions.to_dict(orient='records')
raw_transactions_json = json.dumps(raw_transactions_list)

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>EUPHIRE | Analiza Labirynt Rozmów</title>
    <!-- Google Fonts z EUPHIRE Design System oraz nowoczesny Plus Jakarta Sans -->
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Montserrat:wght@500;600;700;800&family=Roboto+Mono:wght@400;500&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --primary-bg: #F3F7F7; /* Frost White z lekkim odcieniem teal */
            --header-bg: linear-gradient(135deg, #004D54 0%, #002E32 100%); /* Evergreen do Nocturne */
            --text-main: #1D2A2B; /* Obsidian Black z odcieniem ciemnego teal */
            --text-header-light: #FAFAFA;
            --accent: #FCAE2F; /* Ember */
            --accent-subtle: #B2CACC; /* Mist */
            --card-shadow: 0 10px 30px rgba(0, 77, 84, 0.05);
            --card-hover-shadow: 0 20px 40px rgba(0, 77, 84, 0.12);
            
            --font-headings: 'Montserrat', sans-serif;
            --font-paragraphs: 'Plus Jakarta Sans', sans-serif;
            --font-data: 'Plus Jakarta Sans', sans-serif;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            background-color: var(--primary-bg);
            color: var(--text-main);
            font-family: var(--font-paragraphs);
            line-height: 1.6;
            min-height: 100vh;
            -webkit-font-smoothing: antialiased;
        }
        
        /* Typography System */
        h1, h2, h3, h4 {
            font-family: var(--font-headings);
            font-weight: 700;
        }
        
        .overline {
            font-family: 'Roboto Mono', monospace;
            line-height: 1.5;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            font-size: 0.8rem;
            opacity: 0.8;
        }
        
        .data-value {
            font-family: var(--font-data);
            font-weight: 800;
            letter-spacing: -0.03em;
        }
        
        .italic-highlight {
            font-style: italic;
            opacity: 0.9;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem 4rem 2rem;
        }
        
        /* HEADER SECTION */
        .top-header {
            background: var(--header-bg);
            padding: 2.5rem 0;
            margin-bottom: 2rem;
            color: var(--text-header-light);
            border-bottom: 5px solid var(--accent);
            position: relative;
            overflow: hidden;
            box-shadow: 0 10px 30px rgba(0, 46, 50, 0.15);
        }
        
        .top-header::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0; bottom: 0;
            background-image: url('data:image/svg+xml;utf8,<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg"><filter id="noiseFilter"><feTurbulence type="fractalNoise" baseFrequency="0.8" numOctaves="3" stitchTiles="stitch"/></filter><rect width="100%" height="100%" filter="url(%23noiseFilter)" opacity="0.035"/></svg>');
            pointer-events: none;
        }
        
        .header-content {
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: relative;
            z-index: 1;
        }
        
        .logo-container img {
            height: 52px;
            display: block;
            filter: drop-shadow(0 2px 4px rgba(0,0,0,0.1));
        }
        
        .header-text h1 {
            font-size: 2.5rem;
            margin-top: 0.25rem;
            letter-spacing: -0.02em;
            font-weight: 800;
        }
        
        .header-text .overline {
            color: var(--accent);
        }
        
        /* FILTER CONTROLS (Floating Glassmorphism Card) */
        .filter-container {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid rgba(178, 202, 204, 0.4);
            border-radius: 12px;
            padding: 1.25rem 2rem;
            margin-bottom: 2.5rem;
            display: flex;
            flex-wrap: wrap;
            gap: 1.5rem;
            align-items: center;
            justify-content: space-between;
            box-shadow: var(--card-shadow);
            transition: all 0.3s ease;
        }
        
        .filter-container:hover {
            border-color: rgba(0, 77, 84, 0.3);
            box-shadow: 0 12px 35px rgba(0, 77, 84, 0.08);
        }
        
        .filter-group {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            flex-wrap: wrap;
        }
        
        .filter-label {
            font-family: var(--font-headings);
            font-weight: 700;
            color: #004D54;
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        
        .filter-input {
            font-family: var(--font-paragraphs);
            font-weight: 500;
            border: 1.5px solid rgba(178, 202, 204, 0.6);
            border-radius: 6px;
            padding: 0.5rem 0.9rem;
            color: var(--text-main);
            outline: none;
            font-size: 0.9rem;
            background: #ffffff;
            transition: all 0.2s ease-in-out;
        }
        
        .filter-input:focus {
            border-color: #004D54;
            box-shadow: 0 0 0 4px rgba(0, 77, 84, 0.1);
        }
        
        .btn-group {
            display: flex;
            gap: 0.4rem;
            flex-wrap: wrap;
        }
        
        .filter-btn {
            background: rgba(178, 202, 204, 0.15);
            border: 1.5px solid transparent;
            color: #004D54;
            font-family: var(--font-headings);
            font-weight: 700;
            font-size: 0.75rem;
            padding: 0.5rem 1rem;
            border-radius: 6px;
            cursor: pointer;
            text-transform: uppercase;
            letter-spacing: 0.03em;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .filter-btn:hover {
            background: rgba(0, 77, 84, 0.1);
            color: #004D54;
        }
        
        .filter-btn.active {
            background: #004D54;
            color: var(--text-header-light);
            box-shadow: 0 4px 12px rgba(0, 77, 84, 0.2);
        }
        
        /* SUMMARY GRID */
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 1.5rem;
            margin-bottom: 3rem;
        }
        
        .card {
            background-color: #ffffff;
            border-radius: 12px;
            padding: 1.75rem;
            border: 1px solid rgba(178, 202, 204, 0.3);
            box-shadow: var(--card-shadow);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
        }
        
        .card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 4px;
            background: #004D54;
            border-top-left-radius: 12px;
            border-top-right-radius: 12px;
        }
        
        .card:hover {
            transform: translateY(-4px);
            box-shadow: var(--card-hover-shadow);
            border-color: rgba(0, 77, 84, 0.2);
            z-index: 10;
        }
        
        .card h3 {
            color: #004D54;
            font-size: 0.95rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 0.75rem;
            font-weight: 700;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .card .value {
            font-size: 2.4rem;
            color: var(--text-main);
            margin-bottom: 0.5rem;
            line-height: 1.1;
        }
        
        .card .value.accent {
            color: #004D54;
            background: linear-gradient(120deg, #004D54 0%, #002E32 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        
        .card p {
            font-size: 0.85rem;
            color: #607274;
            font-weight: 500;
        }

        /* Tooltip styling */
        .info-tooltip-container {
            position: relative;
            display: inline-block;
            cursor: pointer;
            margin-left: 6px;
        }

        .info-icon {
            font-size: 0.75rem;
            color: var(--accent-subtle);
            background: rgba(0, 77, 84, 0.06);
            width: 17px;
            height: 17px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
            font-weight: 700;
            font-family: sans-serif;
            transition: all 0.2s ease;
            border: 1px solid rgba(178, 202, 204, 0.5);
        }

        .info-icon:hover {
            background: var(--accent);
            color: var(--text-main);
            border-color: var(--accent);
        }

        .tooltip-text {
            visibility: hidden;
            width: 300px;
            max-width: 80vw;
            background-color: #002E32;
            color: #fff;
            text-align: left;
            border-radius: 8px;
            padding: 0.85rem 1.1rem;
            position: absolute;
            z-index: 100;
            bottom: 130%;
            right: -10px;
            opacity: 0;
            transition: opacity 0.3s, visibility 0.3s;
            font-size: 0.9rem;
            font-family: var(--font-paragraphs);
            font-weight: 500;
            line-height: 1.5;
            box-shadow: 0 10px 25px rgba(0,0,0,0.2);
            border: 1px solid rgba(252, 174, 47, 0.3);
            text-transform: none;
            letter-spacing: normal;
        }

        .tooltip-text::after {
            content: "";
            position: absolute;
            top: 100%;
            right: 14px;
            border-width: 5px;
            border-style: solid;
            border-color: #002E32 transparent transparent transparent;
        }

        .info-tooltip-container:hover .tooltip-text {
            visibility: visible;
            opacity: 1;
        }
        
        /* VARIANT ANALYSIS */
        .variant-grid {
            display: flex;
            gap: 1.5rem;
            margin-bottom: 3rem;
            flex-wrap: wrap;
        }
        
        .variant-card {
            flex: 1;
            min-width: 300px;
            background: #ffffff;
            color: var(--text-main);
            border-radius: 12px;
            padding: 1.75rem;
            border: 1px solid rgba(178, 202, 204, 0.3);
            box-shadow: var(--card-shadow);
            position: relative;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .variant-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 4px;
            background: linear-gradient(90deg, #004D54 0%, #FCAE2F 100%);
            border-top-left-radius: 12px;
            border-top-right-radius: 12px;
        }
        
        .variant-card:hover {
            transform: translateY(-4px);
            box-shadow: var(--card-hover-shadow);
            border-color: rgba(0, 77, 84, 0.2);
            z-index: 10;
        }
        
        .variant-card h4 {
            color: #004D54;
            font-family: var(--font-headings);
            font-size: 0.95rem;
            margin-bottom: 1.25rem;
            text-transform: uppercase;
            letter-spacing: 0.03em;
            font-weight: 700;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .variant-stats {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            margin-top: 1rem;
        }
        
        .variant-rev { 
            font-size: 1.8rem; 
            font-family: var(--font-data); 
            font-weight: 800; 
            color: var(--text-main);
            letter-spacing: -0.02em;
        }
        
        .variant-count { 
            font-size: 1rem; 
            color: #607274; 
            font-family: var(--font-data);
            font-weight: 600;
        }
 
        /* CHARTS AREA */
        .section-title {
            margin-bottom: 1.5rem;
            font-size: 1.5rem;
            color: #004D54;
            display: flex;
            align-items: center;
            gap: 1rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        
        .section-title::after {
            content: "";
            flex: 1;
            height: 2px;
            background: linear-gradient(90deg, rgba(178, 202, 204, 0.5) 0%, rgba(178, 202, 204, 0.05) 100%);
        }
        
        .charts-grid-row1 {
            display: grid;
            grid-template-columns: 1fr;
            gap: 1.5rem;
            margin-bottom: 2rem;
        }
        
        @media (min-width: 992px) {
            .charts-grid-row1 {
                grid-template-columns: 1fr 1fr;
            }
        }
        
        .charts-grid-row2 {
            display: grid;
            grid-template-columns: 1fr;
            gap: 1.5rem;
            margin-bottom: 3rem;
        }
        
        @media (min-width: 768px) {
            .charts-grid-row2 {
                grid-template-columns: repeat(2, 1fr);
            }
        }
        
        @media (min-width: 1200px) {
            .charts-grid-row2 {
                grid-template-columns: repeat(3, 1fr);
            }
        }
 
        .chart-container {
            background-color: #ffffff;
            border-radius: 12px;
            padding: 1.75rem;
            border: 1px solid rgba(178, 202, 204, 0.3);
            box-shadow: var(--card-shadow);
            height: 400px;
            transition: all 0.3s ease;
        }
        
        .chart-container:hover {
            box-shadow: 0 12px 35px rgba(0, 77, 84, 0.06);
        }
        
        /* INSIGHTS */
        .insights {
            background-color: #ffffff;
            border-radius: 12px;
            padding: 2.25rem;
            border: 1px solid rgba(178, 202, 204, 0.3);
            border-left: 6px solid #004D54;
            margin-bottom: 2rem;
            box-shadow: var(--card-shadow);
            transition: all 0.3s ease;
        }
        
        .insights:hover {
            box-shadow: var(--card-hover-shadow);
        }
        
        .insights h2 {
            color: #004D54;
            margin-bottom: 1.5rem;
            font-size: 1.5rem;
            letter-spacing: -0.01em;
        }
        
        .insights ul {
            list-style: none;
            margin-top: 1rem;
        }
        
        .insights li {
            margin-bottom: 1.25rem;
            padding-left: 1.75rem;
            position: relative;
            font-size: 1.05rem;
            color: #2F3E40;
        }
        
        .insights li::before {
            content: "";
            width: 8px;
            height: 8px;
            background: var(--accent);
            border-radius: 50%;
            position: absolute;
            left: 0.25rem;
            top: 0.65rem;
        }
        
        .badge {
            background: rgba(0, 77, 84, 0.08);
            color: #004D54;
            padding: 3px 10px;
            border-radius: 6px;
            font-family: 'Roboto Mono', monospace;
            font-size: 0.85rem;
            font-weight: 600;
        }
        
        .badge-rev {
            font-family: var(--font-data);
            font-weight: 700;
            color: #004D54;
        }
        
        /* RESPONSIVE DESIGN - MOBILE FIXES */
        @media (max-width: 768px) {
            .container {
                padding: 0 1rem 2rem 1rem;
            }
            .header-content {
                flex-direction: column;
                text-align: center;
                gap: 1.5rem;
                padding: 0 1rem;
            }
            .header-text h1 {
                font-size: 1.8rem;
            }
            .filter-container {
                padding: 1rem;
                flex-direction: column;
                align-items: stretch;
                gap: 1rem;
            }
            .filter-group {
                flex-direction: column;
                align-items: stretch;
                gap: 0.5rem;
            }
            .filter-input {
                width: 100%;
            }
            .btn-group {
                justify-content: center;
                width: 100%;
            }
            .filter-btn {
                flex: 1;
                text-align: center;
                padding: 0.5rem 0.5rem;
                font-size: 0.7rem;
            }
            .card {
                padding: 1.25rem;
            }
            .card .value {
                font-size: 1.8rem;
            }
            .variant-card {
                padding: 1.25rem;
                min-width: 100%;
            }
            .variant-rev {
                font-size: 1.5rem;
            }
            .chart-container {
                padding: 1rem;
                height: 320px !important;
            }
            .chart-container h3 {
                font-size: 0.95rem !important;
            }
            .chart-container > div {
                height: 230px !important;
            }
        }
    </style>
</head>
<body>
    <div class="top-header">
        <div class="header-content">
            <div class="header-text">
                <span class="overline">Raport Analityczny / Produkt</span>
                <h1>Sprzedaż: Labirynt Rozmów</h1>
                <span class="overline">Wygenerowano: {{GENERATED_DATE}}</span>
            </div>
            <div class="logo-container">
                <img src="https://euphire.pl/wp-content/uploads/2025/11/logo_primary.svg" alt="Euphire Logo">
            </div>
        </div>
    </div>
 
    <div class="container">
        
        <!-- DYNAMIC FILTER CONTAINER -->
        <div class="filter-container">
            <div class="filter-group">
                <span class="filter-label">Wariant:</span>
                <select id="variantSelect" class="filter-input" style="cursor: pointer; font-weight: 600; padding-right: 1.5rem;">
                    <option value="all">Wszystkie warianty</option>
                    <option value="Wersja Fizyczna (Pudełko)">Wersja Fizyczna (Pudełko)</option>
                    <option value="Wersja Cyfrową (On-line)">Wersja Cyfrową (On-line)</option>
                    <option value="Pakiet (Pudełko + On-line)">Pakiet (Pudełko + On-line)</option>
                </select>
            </div>
            
            <div class="filter-group">
                <span class="filter-label">Zakres dat:</span>
                <input type="date" id="startDate" class="filter-input">
                <span class="filter-label" style="font-weight: normal; color: #555;">do</span>
                <input type="date" id="endDate" class="filter-input">
            </div>
            <div class="btn-group">
                <button class="filter-btn" id="btn-14d">Ostatnie 14 dni</button>
                <button class="filter-btn" id="btn-30d">Ostatnie 30 dni</button>
                <button class="filter-btn" id="btn-90d">Ostatnie 90 dni</button>
                <button class="filter-btn active" id="btn-all">Cały okres</button>
            </div>
        </div>
        
        <h2 class="section-title">Podsumowanie Okresu</h2>
        <div class="summary-grid">
            <div class="card">
                <h3>Całkowity Przychód
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Łączna kwota brutto, która wpłynęła od klientów po odliczeniu wszystkich zniżek. Pokazuje realne pieniądze zarobione na sprzedaży.</span>
                    </span>
                </h3>
                <div class="value accent data-value" id="total-revenue">0,00 zł</div>
                <p>Uwzględnia faktyczne wpłaty (po zniżkach)</p>
            </div>
            <div class="card">
                <h3>Liczba Zamówień
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Łączna liczba sfinalizowanych zakupów. Każde zamówienie oznacza transakcję jednego klienta w sklepie.</span>
                    </span>
                </h3>
                <div class="value data-value" id="total-orders">0</div>
                <p>Potwierdzone transakcje w zbiorze</p>
            </div>
            <div class="card">
                <h3>Łącznie Sprzedano (Sztuk)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Suma sprzedanych fizycznych i cyfrowych sztuk gry. Jeśli klient kupi więcej niż jedną sztukę w zamówieniu, ta wartość będzie większa od liczby zamówień.</span>
                    </span>
                </h3>
                <div class="value data-value" id="total-sales">0</div>
                <p>Obejmuje sztuki wszystkich wariantów</p>
            </div>
            <div class="card">
                <h3>Średni Koszyk (AOV)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Średnia wartość jednego zamówienia (czyli ile przeciętnie klient wydaje przy jednym zakupie). Wyliczana jako całkowity przychód podzielony przez liczbę zamówień.</span>
                    </span>
                </h3>
                <div class="value data-value" id="avg-order-value">0,00 zł</div>
                <p>Średni koszyk z zakupem wariantu</p>
            </div>
        </div>
        
        <h2 class="section-title" style="color: #004D54; border-color: rgba(252, 174, 47, 0.4);">Analiza Użycia Rabatów & Marży</h2>
        <div class="summary-grid" style="grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); margin-bottom: 3rem;">
            <div class="card" style="border-left: 4px solid var(--accent);">
                <h3>Udział Zamówień z Kodem
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Procent transakcji, w których klient użył kodu rabatowego. Pokazuje, jak duża część sprzedaży opiera się na promocjach.</span>
                    </span>
                </h3>
                <div class="value data-value" id="discount-pct">0%</div>
                <p id="discount-count-ratio">0 z 0 zamówień</p>
            </div>
            <div class="card" style="border-left: 4px solid var(--accent);">
                <h3>AOV: Bez Kodu vs Z Kodem
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Porównanie średniej kwoty zamówienia dla klientów płacących pełną cenę i tych korzystających ze zniżek. Pomaga ocenić wpływ rabatów na wielkość zakupów.</span>
                    </span>
                </h3>
                <div class="value data-value" style="font-size: 1.25rem; display: flex; flex-direction: column; gap: 0.2rem; margin-top: 0.2rem; font-family: var(--font-data);" id="aov-comparison">
                    <span style="color: #004D54; font-weight: 700;">Bez kodu: <strong id="aov-no-code" style="font-size: 1.35rem;">0,00 zł</strong></span>
                    <span style="color: var(--accent); font-weight: 700;">Z kodem: <strong id="aov-with-code" style="font-size: 1.35rem;">0,00 zł</strong></span>
                </div>
                <p>Średni koszyk z kodem vs bez kodu</p>
            </div>
            <div class="card" style="border-left: 4px solid var(--accent);">
                <h3>Przychód z Rabatów
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Suma pieniędzy, która wpłynęła do sklepu z zamówień promocyjnych (czyli od klientów, którzy wpisali kod rabatowy).</span>
                    </span>
                </h3>
                <div class="value data-value" id="discount-revenue">0,00 zł</div>
                <p>Suma wpłat od osób używających kodów</p>
            </div>
            <div class="card" style="border-left: 4px solid var(--accent);">
                <h3>Udzielone Zniżki (Suma)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Łączna kwota rabatów przyznanych klientom. Pokazuje, ile potencjalnego przychodu 'oddaliśmy' w formie zniżek na rzecz promocji.</span>
                    </span>
                </h3>
                <div class="value data-value" id="discount-saved">0,00 zł</div>
                <p>Kwota zaoszczędzona przez kupujących</p>
            </div>
        </div>
        
        <h2 class="section-title">Wyniki Koszykowe Na Warianty</h2>
        <div class="variant-grid">
            <div class="variant-card" id="card-fizyczna">
                <h4>Wersja Fizyczna (Pudełko)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Wszystkie wskaźniki (kwota i sztuki) dotyczące wyłącznie tradycyjnego, fizycznego wydania gry w pudełku.</span>
                    </span>
                </h4>
                <div class="variant-stats">
                    <div class="variant-rev" id="rev-fizyczna">0,00 zł</div>
                    <div class="variant-count" id="count-fizyczna">0 szt.</div>
                </div>
            </div>
            <div class="variant-card" id="card-cyfrowa">
                <h4>Wersja Cyfrowa (On-line)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Statystyki dotyczące gry w wersji cyfrowej (dostęp do rozgrywki na urządzeniach online).</span>
                    </span>
                </h4>
                <div class="variant-stats">
                    <div class="variant-rev" id="rev-cyfrowa">0,00 zł</div>
                    <div class="variant-count" id="count-cyfrowa">0 szt.</div>
                </div>
            </div>
            <div class="variant-card" id="card-pakiet">
                <h4>Pakiet (Pudełko + On-line)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Statystyki zakupu pakietu hybrydowego: pudełko fizyczne z jednoczesnym dostępem online.</span>
                    </span>
                </h4>
                <div class="variant-stats">
                    <div class="variant-rev" id="rev-pakiet">0,00 zł</div>
                    <div class="variant-count" id="count-pakiet">0 szt.</div>
                </div>
            </div>
        </div>
        
        <h2 class="section-title" style="color: #004D54; border-color: rgba(252, 174, 47, 0.45);">Droga Do Sukcesu</h2>
        <div style="margin-bottom: 3rem;">
            <div class="chart-container" style="height: 400px; width: 100%;">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: #004D54; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.03em; display: flex; align-items: center; justify-content: space-between;">
                    Skumulowany Przychód w Czasie (Wykres Satysfakcji)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Ten wykres pokazuje, jak przychód rósł i kumulował się z dnia na dzień. Linia stale pnie się w górę, obrazując całkowitą zarobioną kwotę od początku wybranego okresu.</span>
                    </span>
                </h3>
                <div style="position: relative; height: 330px; width: 100%;">
                    <canvas id="cumulativeChart"></canvas>
                </div>
            </div>
        </div>
        
        <h2 class="section-title">Dynamika Sprzedaży</h2>
        <div class="charts-grid-row1">
            <div class="chart-container" style="height: 380px;">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: #004D54; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.03em; display: flex; align-items: center; justify-content: space-between;">
                    Trendy Przychodu (Dzienne)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Dzienna wysokość sprzedaży. Pozwala dostrzec nagłe skoki (np. w dniach kampanii mailingowych lub reklam) oraz naturalne wahania.</span>
                    </span>
                </h3>
                <div style="position: relative; height: 310px; width: 100%;">
                    <canvas id="dailyChart"></canvas>
                </div>
            </div>
            <div class="chart-container" style="height: 380px;">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: #004D54; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.03em; display: flex; align-items: center; justify-content: space-between;">
                    Przychód Miesięczny (MoM)
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Ujęcie miesięczne pokazujące łączną sprzedaż w kolejnych miesiącach. Idealne do porównywania wzrostu miesiąc do miesiąca.</span>
                    </span>
                </h3>
                <div style="position: relative; height: 310px; width: 100%;">
                    <canvas id="monthlyChart"></canvas>
                </div>
            </div>
        </div>
        
        <h2 class="section-title">Analiza Marketingowa & Kampanie</h2>
        <div class="charts-grid-row2">
            <div class="chart-container">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: #004D54; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.03em; display: flex; align-items: center; justify-content: space-between;">
                    Przychód Tygodniowy
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Suma przychodów pogrupowana w cykle tygodniowe. Ułatwia analizę średnioterminowych trendów bez dziennych zakłóceń.</span>
                    </span>
                </h3>
                <div style="position: relative; height: 310px; width: 100%;">
                    <canvas id="weeklyChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: #004D54; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.03em; display: flex; align-items: center; justify-content: space-between;">
                    Sprzedaż wg Dni Tygodnia
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Łączna sprzedaż z podziałem na dni tygodnia. Pokazuje, w które dni klienci kupują najchętniej (np. weekendy vs środek tygodnia), co pomaga planować start reklam.</span>
                    </span>
                </h3>
                <div style="position: relative; height: 310px; width: 100%;">
                    <canvas id="dayOfWeekChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: #004D54; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.03em; display: flex; align-items: center; justify-content: space-between;">
                    Przychody: Rabat vs Standard
                    <span class="info-tooltip-container">
                        <span class="info-icon">i</span>
                        <span class="tooltip-text">Podział całkowitego przychodu na kwoty uzyskane ze sprzedaży w standardowej cenie oraz ze sprzedaży z użyciem kodów rabatowych.</span>
                    </span>
                </h3>
                <div style="position: relative; height: 300px; width: 100%; display: flex; justify-content: center; align-items: center;">
                    <canvas id="discountChart" style="max-height: 260px; max-width: 260px;"></canvas>
                </div>
            </div>
        </div>
        
        <div class="insights">
            <h2>Wnioski Zgodne z <span class="italic-highlight">Faktami, Metodą, Praktyką</span></h2>
            <p>Zautomatyzowana analiza odnotowała następujące momenty skrajne z podzbioru danych w wybranym przedziale czasowym:</p>
            <ul>
                <li><strong>Najsilniejszy dzień pod względem przychodów:</strong> Dzień <span class="badge" id="best-day-badge">-</span> z wynikiem <span class="badge-rev" id="best-day-rev">0,00 zł</span>.</li>
                <li><strong>Najsilniejszy tydzień łączny:</strong> <span class="badge" id="best-week-badge">-</span> generujący przychód w wysokości <span class="badge-rev" id="best-week-rev">0,00 zł</span>.</li>
                <li>Zestawienie <span class="italic-highlight">wariantów pudełkowych oraz online</span> wskazuje, jak użytkownicy wybierają formę gry w skali danego ujęcia czasu. Pamiętaj, aby śledzić spadek / przyrost zainteresowania pakietami premium w okresach przedświątecznych lub po kampaniach mailingowych marki Euphire.</li>
            </ul>
        </div>
    </div>
 
    <script>
        const rawTransactions = {{RAW_TRANSACTIONS_JSON}};
        
        function getISOWeek(dateString) {
            const date = new Date(dateString);
            const tempDate = new Date(date.valueOf());
            tempDate.setDate(tempDate.getDate() + 4 - (tempDate.getDay() || 7));
            const yearStart = new Date(tempDate.getFullYear(), 0, 1);
            const weekNo = Math.ceil((((tempDate - yearStart) / 86400000) + 1) / 7);
            const weekStr = weekNo < 10 ? '0' + weekNo : weekNo;
            return `${tempDate.getFullYear()}-W${weekStr}`;
        }
 
        Chart.defaults.font.family = "'Plus Jakarta Sans', sans-serif";
        Chart.defaults.color = "#1D2A2B";
 
        // Konfiguracja tooltipów w formacie waluty PLN (ładne formatowanie dla człowieka)
        const tooltipCurrencyCallback = {
            callbacks: {
                label: function(context) {
                    let label = context.label || context.dataset.label || '';
                    if (label) {
                        label += ': ';
                    }
                    let val = 0;
                    if (context.parsed !== undefined && context.parsed.y !== undefined) {
                        val = context.parsed.y;
                    } else if (context.parsed !== undefined && typeof context.parsed === 'number') {
                        val = context.parsed;
                    } else {
                        val = context.raw || 0;
                    }
                    label += val.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
                    return label;
                }
            }
        };

        // Wspólne opcje osi Y z ładnym formatowaniem zł
        const currencyYScale = {
            y: {
                beginAtZero: true,
                ticks: {
                    callback: function(value) {
                        return value.toLocaleString('pl-PL') + ' zł';
                    }
                }
            }
        };

        // Ograniczenie gęstości osi X
        const cleanXScale = {
            x: {
                ticks: {
                    maxTicksLimit: 12,
                    maxRotation: 0,
                    autoSkip: true
                },
                grid: {
                    display: false
                }
            }
        };
 
        const ctxDaily = document.getElementById('dailyChart').getContext('2d');
        const dailyChartInstance = new Chart(ctxDaily, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Przychód Dzienny', data: [], borderColor: '#004D54', borderWidth: 2.5, pointRadius: 1, fill: true, backgroundColor: 'rgba(0, 77, 84, 0.07)', tension: 0.3 }] },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { display: false }, tooltip: tooltipCurrencyCallback }, 
                scales: Object.assign({}, currencyYScale, cleanXScale)
            }
        });

        const ctxCumulative = document.getElementById('cumulativeChart').getContext('2d');
        const cumulativeChartInstance = new Chart(ctxCumulative, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Suma Skumulowana', data: [], borderColor: '#FCAE2F', borderWidth: 4, pointRadius: 0, fill: true, backgroundColor: 'rgba(252, 174, 47, 0.1)', tension: 0.3 }] },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false }, tooltip: tooltipCurrencyCallback },
                scales: Object.assign({}, currencyYScale, cleanXScale)
            }
        });
        
        const ctxMonthly = document.getElementById('monthlyChart').getContext('2d');
        const monthlyChartInstance = new Chart(ctxMonthly, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'Przychód', data: [], backgroundColor: '#004D54', borderRadius: 4 }] },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { display: false }, tooltip: tooltipCurrencyCallback }, 
                scales: currencyYScale
            }
        });
        
        const ctxWeekly = document.getElementById('weeklyChart').getContext('2d');
        const weeklyChartInstance = new Chart(ctxWeekly, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'Przychód za Tydzień', data: [], backgroundColor: '#FCAE2F', borderRadius: 4 }] },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { display: false }, tooltip: tooltipCurrencyCallback }, 
                scales: currencyYScale
            }
        });
 
        const ctxDayOfWeek = document.getElementById('dayOfWeekChart').getContext('2d');
        const dayOfWeekChartInstance = new Chart(ctxDayOfWeek, {
            type: 'bar',
            data: {
                labels: ['Pon', 'Wt', 'Śr', 'Czw', 'Pt', 'Sob', 'Niedz'],
                datasets: [{ label: 'Przychód Łączny', data: [0, 0, 0, 0, 0, 0, 0], backgroundColor: '#004D54', borderRadius: 4 }]
            },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { display: false }, tooltip: tooltipCurrencyCallback }, 
                scales: currencyYScale
            }
        });
 
        const ctxDiscount = document.getElementById('discountChart').getContext('2d');
        const discountChartInstance = new Chart(ctxDiscount, {
            type: 'doughnut',
            data: { labels: ['Cena standardowa', 'Z rabatem/kodem'], datasets: [{ data: [0, 0], backgroundColor: ['#004D54', '#FCAE2F'], borderWidth: 2 }] },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { position: 'bottom' }, tooltip: tooltipCurrencyCallback }, 
                cutout: '60%' 
            }
        });
 
        // Standardowe punkty cenowe (bez promocji) do wykrywania kodów rabatowych
        const standardPrices = {
            'Wersja Fizyczna (Pudełko)': [99, 109, 120, 129, 139],
            'Wersja Cyfrową (On-line)': [59, 89],
            'Pakiet (Pudełko + On-line)': [169, 199]
        };
 
        function isDiscounted(t) {
            const price = parseFloat(t['Kwota Brutto']) || 0;
            const qty = parseInt(t['Sztuk']) || 1;
            if (qty <= 0) return false;
            const pricePerUnit = price / qty;
            const stdPrices = standardPrices[t.Wariant];
            if (!stdPrices) return false;
            return !stdPrices.some(p => Math.abs(p - pricePerUnit) < 0.2);
        }
        
        const monthNames = {
            '01': 'Styczeń', '02': 'Luty', '03': 'Marzec', '04': 'Kwiecień', '05': 'Maj', '06': 'Czerwiec',
            '07': 'Lipiec', '08': 'Sierpień', '09': 'Wrzesień', '10': 'Październik', '11': 'Listopad', '12': 'Grudzień'
        };
        function formatMonth(yymm) {
            const parts = yymm.split('-');
            const y = parts[0];
            const m = monthNames[parts[1]] || parts[1];
            return `${m} ${y}`;
        }
 
        function updateDashboard(startStr, endStr) {
            const startDate = new Date(startStr + 'T00:00:00');
            const endDate = new Date(endStr + 'T23:59:59');
            
            const selectedVariant = document.getElementById('variantSelect').value;
            
            const filtered = rawTransactions.filter(t => {
                const d = new Date(t.Dzien + 'T00:00:00');
                const matchesDate = d >= startDate && d <= endDate;
                const matchesVariant = selectedVariant === 'all' || t.Wariant === selectedVariant;
                return matchesDate && matchesVariant;
            });
            
            let totalRevenue = 0, totalSales = 0, totalOrders = filtered.length;
            const varStats = {'Wersja Fizyczna (Pudełko)': { rev: 0, count: 0 }, 'Wersja Cyfrową (On-line)': { rev: 0, count: 0 }, 'Pakiet (Pudełko + On-line)': { rev: 0, count: 0 }};
            const dailyMap = {}, weeklyMap = {}, monthlyMap = {}, dayOfWeekMap = [0, 0, 0, 0, 0, 0, 0];
            
            // Kody rabatowe variables
            let discountCount = 0;
            let fullPriceCount = 0;
            let discountRevenue = 0;
            let fullPriceRevenue = 0;
            let totalDiscountSaved = 0;
            
            filtered.forEach(t => {
                const rev = parseFloat(t['Kwota Brutto']) || 0;
                const qty = parseInt(t['Sztuk']) || 0;
                totalRevenue += rev; totalSales += qty;
                if (varStats[t.Wariant]) { varStats[t.Wariant].rev += rev; varStats[t.Wariant].count += qty; }
                
                dailyMap[t.Dzien] = (dailyMap[t.Dzien] || 0) + rev;
                const week = getISOWeek(t.Dzien); weeklyMap[week] = (weeklyMap[week] || 0) + rev;
                const month = t.Dzien.substring(0, 7); monthlyMap[month] = (monthlyMap[month] || 0) + rev;
                
                const dateObj = new Date(t.Dzien + 'T00:00:00');
                const rawDay = dateObj.getDay();
                dayOfWeekMap[rawDay === 0 ? 6 : rawDay - 1] += rev;
                
                // Sprawdzenie rabatowości transakcji
                const isDisc = isDiscounted(t);
                if (isDisc) {
                    discountCount++;
                    discountRevenue += rev;
                    
                    const pricePerUnit = rev / (qty || 1);
                    const stds = standardPrices[t.Wariant] || [];
                    const validStds = stds.filter(s => s >= pricePerUnit);
                    const origUnit = validStds.length > 0 ? Math.min(...validStds) : (stds.length > 0 ? Math.max(...stds) : pricePerUnit);
                    totalDiscountSaved += (origUnit - pricePerUnit) * qty;
                } else {
                    fullPriceCount++;
                    fullPriceRevenue += rev;
                }
            });
            
            // Elementy podsumowania okresu
            document.getElementById('total-revenue').innerText = totalRevenue.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('total-orders').innerText = totalOrders.toLocaleString('pl-PL');
            document.getElementById('total-sales').innerText = totalSales.toLocaleString('pl-PL');
            document.getElementById('avg-order-value').innerText = (totalOrders > 0 ? (totalRevenue / totalOrders) : 0).toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            
            // Kody rabatowe panel
            const discountPct = totalOrders > 0 ? (discountCount / totalOrders * 100) : 0;
            document.getElementById('discount-pct').innerText = discountPct.toFixed(1) + '%';
            document.getElementById('discount-count-ratio').innerText = `${discountCount} z ${totalOrders} zamówień`;
            
            const aovNoCode = fullPriceCount > 0 ? (fullPriceRevenue / fullPriceCount) : 0;
            const aovWithCode = discountCount > 0 ? (discountRevenue / discountCount) : 0;
            document.getElementById('aov-no-code').innerText = aovNoCode.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('aov-with-code').innerText = aovWithCode.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            
            document.getElementById('discount-revenue').innerText = discountRevenue.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('discount-saved').innerText = totalDiscountSaved.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            
            // Warianty pudełkowe/cyfrowe
            document.getElementById('rev-fizyczna').innerText = varStats['Wersja Fizyczna (Pudełko)'].rev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('count-fizyczna').innerText = varStats['Wersja Fizyczna (Pudełko)'].count.toLocaleString('pl-PL') + ' szt.';
            document.getElementById('rev-cyfrowa').innerText = varStats['Wersja Cyfrową (On-line)'].rev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('count-cyfrowa').innerText = varStats['Wersja Cyfrową (On-line)'].count.toLocaleString('pl-PL') + ' szt.';
            document.getElementById('rev-pakiet').innerText = varStats['Pakiet (Pudełko + On-line)'].rev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('count-pakiet').innerText = varStats['Pakiet (Pudełko + On-line)'].count.toLocaleString('pl-PL') + ' szt.';
            
            // Najsilniejsze okresy (Insights)
            let bestDay = '-', bestDayRev = 0;
            Object.keys(dailyMap).forEach(d => {
                if (dailyMap[d] > bestDayRev) {
                    bestDay = d;
                    bestDayRev = dailyMap[d];
                }
            });
            
            let bestWeek = '-', bestWeekRev = 0;
            Object.keys(weeklyMap).forEach(w => {
                if (weeklyMap[w] > bestWeekRev) {
                    bestWeek = w;
                    bestWeekRev = weeklyMap[w];
                }
            });
            
            document.getElementById('best-day-badge').innerText = bestDay !== '-' ? bestDay : '-';
            document.getElementById('best-day-rev').innerText = bestDayRev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('best-week-badge').innerText = bestWeek !== '-' ? bestWeek : '-';
            document.getElementById('best-week-rev').innerText = bestWeekRev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
 
            // Aktualizacja wykresów
            const sortedDays = Object.keys(dailyMap).sort();
            dailyChartInstance.data.labels = sortedDays;
            dailyChartInstance.data.datasets[0].data = sortedDays.map(k => dailyMap[k]);
            dailyChartInstance.update();

            // Przychód skumulowany (Suma skumulowana)
            let runningSum = 0;
            const cumulativeData = sortedDays.map(d => {
                runningSum += dailyMap[d];
                return runningSum;
            });
            cumulativeChartInstance.data.labels = sortedDays;
            cumulativeChartInstance.data.datasets[0].data = cumulativeData;
            cumulativeChartInstance.update();
            
            const sortedMonths = Object.keys(monthlyMap).sort();
            monthlyChartInstance.data.labels = sortedMonths.map(formatMonth);
            monthlyChartInstance.data.datasets[0].data = sortedMonths.map(k => monthlyMap[k]);
            monthlyChartInstance.update();
            
            weeklyChartInstance.data.labels = Object.keys(weeklyMap).sort();
            weeklyChartInstance.data.datasets[0].data = weeklyChartInstance.data.labels.map(k => weeklyMap[k]);
            weeklyChartInstance.update();
            
            dayOfWeekChartInstance.data.datasets[0].data = dayOfWeekMap;
            dayOfWeekChartInstance.update();
            
            discountChartInstance.data.datasets[0].data = [fullPriceRevenue, discountRevenue];
            discountChartInstance.update();
        }
 
        const startInput = document.getElementById('startDate'), endInput = document.getElementById('endDate');
        const sortedDates = rawTransactions.map(t => t.Dzien).sort();
        const minDateStr = sortedDates[0], maxDateStr = sortedDates[sortedDates.length - 1];
        startInput.value = minDateStr; endInput.value = maxDateStr;
 
        function clearButtonActive() { document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.remove('active')); }
        ['14d', '30d', '90d'].forEach(id => document.getElementById('btn-' + id).addEventListener('click', (e) => {
            clearButtonActive(); e.target.classList.add('active');
            let end = new Date(maxDateStr + 'T00:00:00'), start = new Date(maxDateStr + 'T00:00:00');
            start.setDate(end.getDate() - parseInt(id));
            const startStr = start.toISOString().split('T')[0];
            startInput.value = startStr; endInput.value = maxDateStr;
            updateDashboard(startStr, maxDateStr);
        }));
        document.getElementById('btn-all').addEventListener('click', (e) => { clearButtonActive(); e.target.classList.add('active'); startInput.value = minDateStr; endInput.value = maxDateStr; updateDashboard(minDateStr, maxDateStr); });
        
        startInput.addEventListener('change', () => { clearButtonActive(); updateDashboard(startInput.value, endInput.value); });
        endInput.addEventListener('change', () => { clearButtonActive(); updateDashboard(startInput.value, endInput.value); });
        document.getElementById('variantSelect').addEventListener('change', () => { updateDashboard(startInput.value, endInput.value); });
        
        updateDashboard(minDateStr, maxDateStr);
    </script>
</body>
</html>
"""

# Przeprowadź zamianę placeholderów w szablonie HTML
html_content = HTML_TEMPLATE.replace('{{GENERATED_DATE}}', datetime.now().strftime('%d.%m.%Y %H:%M'))
html_content = html_content.replace('{{RAW_TRANSACTIONS_JSON}}', raw_transactions_json)

# Zapisz do index.html (główna strona dla GitHub Pages) oraz Raport_Labirynt_Euphire.html
output_paths = ['index.html', 'Raport_Labirynt_Euphire.html']
for opath in output_paths:
    with open(opath, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"Raport wygenerowany: {os.path.abspath(opath)}")

# --- AUTOMATYZACJA GIT & GITHUB PAGES ---
try:
    # Sprawdź czy jesteśmy w repozytorium git
    git_status = subprocess.run(['git', 'rev-parse', '--is-inside-work-tree'], capture_output=True, text=True)
    if git_status.returncode == 0:
        print("\nWykryto repozytorium Git. Publikuję raport na GitHub Pages...")
        
        # 1. Dodaj pliki do indeksu
        subprocess.run(['git', 'add', 'index.html', 'Raport_Labirynt_Euphire.html', '.gitignore', 'generate_report.py', 'Specyfikacja_Analityczna_Labirynt.md', 'README.md', 'Specyfikacja_Wdrozeniowa_Agenta.md'], check=True)
        
        # 2. Wykonaj commit (jeśli są zmiany)
        commit_check = subprocess.run(['git', 'status', '--porcelain'], capture_output=True, text=True)
        if commit_check.stdout.strip():
            commit_msg = f"Update report with currency tooltips, info tooltips and cumulative revenue chart - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            subprocess.run(['git', 'commit', '-m', commit_msg], check=True)
            print("Zatwierdzono nowe zmiany w Git.")
            
            # 3. Pchnij zmiany na serwer (jeśli remote istnieje)
            remote_check = subprocess.run(['git', 'remote', 'get-url', 'origin'], capture_output=True, text=True)
            if remote_check.returncode == 0:
                push_res = subprocess.run(['git', 'push', 'origin', 'main'], capture_output=True, text=True)
                if push_res.returncode == 0:
                    print("[SUKCES] Zmiany zostały wysłane do repozytorium GitHub!")
                    print(f"Link do raportu dla marketera: https://baciok91.github.io/euphire-raporty-sprzedazy/")
                else:
                    print(f"[BŁĄD] Nie udało się wysłać zmian do GitHub origin: {push_res.stderr.strip()}")
            else:
                print("[INFO] Brak zdalnego repozytorium origin. Skonfiguruj je, aby automatycznie publikować.")
        else:
            print("Brak nowych zmian w raporcie do wysłania.")
except Exception as e:
    print(f"Błąd podczas automatyzacji Git: {e}")
