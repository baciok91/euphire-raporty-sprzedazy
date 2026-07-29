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
    excel_files = glob.glob('*am*wienia*.xlsx')
    if not excel_files:
        excel_files = [f for f in glob.glob('*.xlsx') if f != 'od wrzesnia do maja.xlsx' and not f.startswith('~$')]
    
    if excel_files:
        excel_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        file_path = excel_files[0]
    else:
        file_path = 'od wrzesnia do maja.xlsx'

print(f"Używam pliku bazy danych: {file_path}")

# Wczytaj dane
df = pd.read_excel(file_path, header=2)

# Oczyść wiersze z uszkodzoną datą i nagłówki podsumowujące
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
df = df.dropna(subset=['Data']).copy()
df['Kwota Brutto'] = pd.to_numeric(df['Kwota Brutto'], errors='coerce').fillna(0)
df['Sztuk'] = pd.to_numeric(df['Sztuk'], errors='coerce').fillna(0)

# Odrzuć wiersze podsumowujące
df = df[~df['Nazwa Kursu'].astype(str).str.contains('23%|5%|Suma|Paragony|Faktury', na=False)].copy()

TARGET_LABIRYNT = {
    'Labirynt Rozmów | Gra Psychologiczna dla Par i Przyjaciół [Gra karciana]': 'Wersja Fizyczna (Pudełko)',
    'Labirynt Rozmów [GRA ON-LINE]': 'Wersja Cyfrową (On-line)',
    'Labirynt Rozmów [Pudełko Premium + GRA ON-LINE]': 'Pakiet (Pudełko + On-line)'
}

def extract_author(name_str):
    if 'Woydyłło' in name_str:
        return 'dr Ewa Woydyłło-Osiatyńska'
    elif 'de Barbaro' in name_str:
        return 'Prof. Bogdan de Barbaro'
    elif 'Bralczyk' in name_str:
        return 'Prof. Jerzy Bralczyk'
    elif 'Mosak' in name_str:
        return 'Piotr Mosak'
    elif 'Walkiewicz' in name_str:
        return 'Jacek Walkiewicz'
    elif 'Kołodziejczyk' in name_str:
        return 'Maciej Kołodziejczyk'
    elif 'Matejuk' in name_str:
        return 'Piotr Matejuk'
    elif 'Dobiała' in name_str:
        return 'Ewa Dobiała'
    elif 'Piotrowski' in name_str:
        return 'dr Krzysztof Piotrowski'
    elif 'Dębski' in name_str:
        return 'dr Maciej Dębski'
    elif 'Lipiarz' in name_str:
        return 'Adrianna Lipiarz'
    elif 'Slowdating' in name_str:
        return 'Slowdating'
    elif 'PAKIET' in name_str or 'Kompendium' in name_str:
        return 'Pakiety Wieloautorskie'
    else:
        return 'Pozostali Wykładowcy'

records = []
for _, row in df.iterrows():
    name = str(row['Nazwa Kursu']).strip()
    kwota = round(float(row['Kwota Brutto']), 2)
    sztuk = int(row['Sztuk'])
    dt = row['Data']
    dzien = dt.strftime('%Y-%m-%d')
    godzina = dt.strftime('%H:%M')
    
    is_lab = name in TARGET_LABIRYNT
    lab_variant = TARGET_LABIRYNT.get(name, '')
    author = extract_author(name)
    is_mosak = 'Mosak' in name
    is_debarbaro_terapeuta = 'Jak być dobrym psychoterapeutą dla samego siebie' in name
    is_walkiewicz = 'Walkiewicz' in name
    
    records.append({
        'd': dzien,
        'g': godzina,
        'n': name,
        'k': kwota,
        's': sztuk,
        'isL': is_lab,
        'vL': lab_variant,
        'a': author,
        'isM': is_mosak,
        'isB': is_debarbaro_terapeuta,
        'isW': is_walkiewicz
    })

raw_transactions_json = json.dumps(records, ensure_ascii=False)
print(f"Dane wczytane pomyślnie. Liczba wszystkich transakcji: {len(records)}")

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>EUPHIRE | Raporty i Analiza Sprzedaży</title>
    <!-- Google Fonts z EUPHIRE Design System -->
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Montserrat:wght@500;600;700;800&family=Roboto+Mono:wght@400;500&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --primary-bg: #F3F7F7;
            --header-bg: linear-gradient(135deg, #004D54 0%, #002E32 100%);
            --text-main: #1D2A2B;
            --text-header-light: #FAFAFA;
            --accent: #FCAE2F;
            --accent-subtle: #B2CACC;
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
            opacity: 0.85;
        }
        
        .data-value {
            font-family: var(--font-data);
            font-weight: 800;
            letter-spacing: -0.03em;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem 4rem 2rem;
        }
        
        /* HEADER SECTION */
        .top-header {
            background: var(--header-bg);
            padding: 2.5rem 0 3.5rem 0;
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
            font-size: 2.3rem;
            margin-top: 0.25rem;
            letter-spacing: -0.02em;
            font-weight: 800;
        }
        
        .header-text .overline {
            color: var(--accent);
        }

        /* TOP NAVIGATION TABS */
        .nav-tabs-wrapper {
            max-width: 1400px;
            margin: -2.2rem auto 2rem auto;
            padding: 0 2rem;
            position: relative;
            z-index: 20;
        }

        .nav-tabs {
            display: flex;
            gap: 0.6rem;
            flex-wrap: wrap;
        }

        .nav-tab-btn {
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1.5px solid rgba(178, 202, 204, 0.6);
            border-radius: 12px;
            padding: 0.85rem 1.4rem;
            font-family: var(--font-headings);
            font-weight: 700;
            font-size: 0.88rem;
            color: #004D54;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
            box-shadow: 0 6px 16px rgba(0, 46, 50, 0.08);
        }

        .nav-tab-btn:hover {
            background: #ffffff;
            color: #002E32;
            transform: translateY(-2px);
            box-shadow: 0 10px 24px rgba(0, 77, 84, 0.12);
        }

        .nav-tab-btn.active {
            background: #004D54;
            color: #ffffff;
            border-color: #004D54;
            box-shadow: 0 8px 24px rgba(0, 77, 84, 0.25);
            position: relative;
        }

        .nav-tab-btn.active .tab-icon {
            transform: scale(1.15);
        }

        .tab-icon {
            font-size: 1.15rem;
            transition: transform 0.2s ease;
        }
        
        /* FILTER CONTROLS */
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
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2.5rem;
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
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 0.75rem;
            font-weight: 700;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .card .value {
            font-size: 2.2rem;
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

        .card .value.gold {
            color: #D97706;
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
            width: 280px;
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
            font-size: 0.85rem;
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
        
        /* VARIANT GRID / LEADERBOARD */
        .variant-grid {
            display: flex;
            gap: 1.5rem;
            margin-bottom: 2.5rem;
            flex-wrap: wrap;
        }
        
        .variant-card {
            flex: 1;
            min-width: 280px;
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
        }
        
        .variant-card h4 {
            color: #004D54;
            font-family: var(--font-headings);
            font-size: 1.05rem;
            margin-bottom: 0.5rem;
        }
        
        .variant-rev {
            font-size: 1.8rem;
            font-weight: 800;
            color: #004D54;
            margin-bottom: 0.5rem;
            font-family: var(--font-data);
        }
        
        .variant-stats {
            display: flex;
            justify-content: space-between;
            font-size: 0.85rem;
            color: #607274;
            border-top: 1px dashed rgba(178, 202, 204, 0.6);
            padding-top: 0.75rem;
            margin-top: 0.75rem;
        }
        
        .section-title {
            font-size: 1.4rem;
            color: #004D54;
            margin: 2rem 0 1.25rem 0;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            letter-spacing: -0.01em;
        }
        
        /* CHARTS GRID */
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
            margin-bottom: 2.5rem;
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
        
        /* TABLE STYLING */
        .table-container {
            background: #ffffff;
            border-radius: 12px;
            padding: 1.75rem;
            border: 1px solid rgba(178, 202, 204, 0.3);
            box-shadow: var(--card-shadow);
            margin-bottom: 2.5rem;
            overflow-x: auto;
        }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
            text-align: left;
        }

        .data-table th {
            background-color: rgba(0, 77, 84, 0.05);
            color: #004D54;
            font-family: var(--font-headings);
            font-weight: 700;
            padding: 0.9rem 1rem;
            border-bottom: 2px solid rgba(0, 77, 84, 0.15);
            text-transform: uppercase;
            font-size: 0.78rem;
            letter-spacing: 0.03em;
        }

        .data-table td {
            padding: 0.85rem 1rem;
            border-bottom: 1px solid rgba(178, 202, 204, 0.3);
            color: var(--text-main);
        }

        .data-table tr:hover {
            background-color: rgba(0, 77, 84, 0.02);
        }

        .badge-tag {
            background: rgba(0, 77, 84, 0.08);
            color: #004D54;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 0.8rem;
            font-weight: 600;
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
            font-size: 1.4rem;
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
            font-size: 1.02rem;
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
        
        .hero-banner {
            background: linear-gradient(135deg, #00383D 0%, #001F22 100%);
            color: #ffffff;
            border-radius: 12px;
            padding: 2rem 2.25rem;
            margin-bottom: 2.5rem;
            border: 1px solid rgba(252, 174, 47, 0.3);
            box-shadow: 0 10px 30px rgba(0, 46, 50, 0.2);
            position: relative;
            overflow: hidden;
        }

        .hero-banner::before {
            content: '';
            position: absolute;
            top: -50%; right: -20%;
            width: 300px; height: 300px;
            background: radial-gradient(circle, rgba(252, 174, 47, 0.15) 0%, transparent 70%);
            pointer-events: none;
        }

        .hero-banner .hero-badge {
            background: var(--accent);
            color: #1D2A2B;
            font-family: 'Roboto Mono', monospace;
            font-size: 0.75rem;
            font-weight: 800;
            padding: 0.3rem 0.8rem;
            border-radius: 4px;
            text-transform: uppercase;
            display: inline-block;
            margin-bottom: 0.75rem;
        }

        .hero-banner h2 {
            font-size: 1.8rem;
            color: #ffffff;
            margin-bottom: 0.5rem;
        }

        .hero-banner p {
            font-size: 1.05rem;
            color: rgba(255, 255, 255, 0.85);
            max-width: 900px;
        }
        
        /* TAB PANES */
        .tab-pane {
            display: none;
        }
        
        .tab-pane.active {
            display: block;
        }

        /* RESPONSIVE DESIGN - MOBILE FIXES */
        @media (max-width: 768px) {
            .container {
                padding: 0 1rem 2rem 1rem;
            }
            .nav-tabs-wrapper {
                padding: 0 1rem;
                margin-top: -1.8rem;
            }
            .nav-tab-btn {
                padding: 0.75rem 0.9rem;
                font-size: 0.8rem;
                flex: 1 1 auto;
                justify-content: center;
            }
            .header-content {
                flex-direction: column;
                text-align: center;
                gap: 1.5rem;
                padding: 0 1rem;
            }
            .header-text h1 {
                font-size: 1.7rem;
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
            .chart-container {
                padding: 1rem;
                height: 320px !important;
            }
        }
    </style>
</head>
<body>
    <div class="top-header">
        <div class="header-content">
            <div class="header-text">
                <span class="overline">System Raportowy EUPHIRE</span>
                <h1>Dashboard Analityczny Sprzedaży</h1>
                <span class="overline">Aktualizacja danych: {{GENERATED_DATE}}</span>
            </div>
            <div class="logo-container">
                <img src="https://euphire.pl/wp-content/uploads/2025/11/logo_primary.svg" alt="Euphire Logo">
            </div>
        </div>
    </div>

    <!-- MAIN NAVIGATION TABS -->
    <div class="nav-tabs-wrapper">
        <div class="nav-tabs">
            <button class="nav-tab-btn active" data-tab="labirynt">
                <span class="tab-icon">🎴</span> Labirynt Rozmów
            </button>
            <button class="nav-tab-btn" data-tab="szkolenia">
                <span class="tab-icon">🎓</span> Szkolenia Ogółem
            </button>
            <button class="nav-tab-btn" data-tab="mosak">
                <span class="tab-icon">👤</span> Piotr Mosak
            </button>
            <button class="nav-tab-btn" data-tab="debarbaro">
                <span class="tab-icon">👨‍⚕️</span> Prof. de Barbaro („Dobry Terapeuta”)
            </button>
            <button class="nav-tab-btn" data-tab="walkiewicz">
                <span class="tab-icon">🎤</span> Jacek Walkiewicz („Pełna MOC”)
            </button>
        </div>
    </div>
 
    <div class="container">
        
        <!-- ========================================== -->
        <!-- TAB 1: LABIRYNT ROZMÓW -->
        <!-- ========================================== -->
        <div class="tab-pane active" id="pane-labirynt">
            
            <div class="filter-container">
                <div class="filter-group">
                    <span class="filter-label">Wariant:</span>
                    <select id="labVariantSelect" class="filter-input" style="cursor: pointer; font-weight: 600; padding-right: 1.5rem;">
                        <option value="all">Wszystkie warianty</option>
                        <option value="Wersja Fizyczna (Pudełko)">Wersja Fizyczna (Pudełko)</option>
                        <option value="Wersja Cyfrową (On-line)">Wersja Cyfrową (On-line)</option>
                        <option value="Pakiet (Pudełko + On-line)">Pakiet (Pudełko + On-line)</option>
                    </select>
                </div>
                
                <div class="filter-group">
                    <span class="filter-label">Zakres dat:</span>
                    <input type="date" id="labStartDate" class="filter-input">
                    <span class="filter-label" style="font-weight: normal; color: #555;">do</span>
                    <input type="date" id="labEndDate" class="filter-input">
                </div>
                <div class="btn-group">
                    <button class="filter-btn" id="lab-btn-14d">Ostatnie 14 dni</button>
                    <button class="filter-btn" id="lab-btn-30d">Ostatnie 30 dni</button>
                    <button class="filter-btn" id="lab-btn-90d">Ostatnie 90 dni</button>
                    <button class="filter-btn active" id="lab-btn-all">Cały okres</button>
                </div>
            </div>
            
            <h2 class="section-title">🎴 Podsumowanie Wyników: Labirynt Rozmów</h2>
            <div class="summary-grid">
                <div class="card">
                    <h3>Całkowity Przychód</h3>
                    <div class="value accent data-value" id="lab-total-revenue">0,00 zł</div>
                    <p>Uwzględnia zniżki i rabaty</p>
                </div>
                <div class="card">
                    <h3>Liczba Zamówień</h3>
                    <div class="value data-value" id="lab-total-orders">0</div>
                    <p>Potwierdzone transakcje</p>
                </div>
                <div class="card">
                    <h3>Sprzedano Sztuk</h3>
                    <div class="value data-value" id="lab-total-sales">0</div>
                    <p>Wszystkie warianty</p>
                </div>
                <div class="card">
                    <h3>Średni Koszyk (AOV)</h3>
                    <div class="value data-value" id="lab-aov">0,00 zł</div>
                    <p>Przeciętne zamówienie</p>
                </div>
                <div class="card">
                    <h3>Najlepszy Dzień</h3>
                    <div class="value data-value" style="font-size: 1.4rem;" id="lab-best-day-badge">-</div>
                    <p id="lab-best-day-rev">0,00 zł</p>
                </div>
                <div class="card">
                    <h3>Najlepszy Tydzień</h3>
                    <div class="value data-value" style="font-size: 1.4rem;" id="lab-best-week-badge">-</div>
                    <p id="lab-best-week-rev">0,00 zł</p>
                </div>
            </div>

            <!-- VARIANT ANALYSIS -->
            <div class="variant-grid">
                <div class="variant-card">
                    <h4>Wersja Fizyczna (Pudełko)</h4>
                    <div class="variant-rev" id="lab-v-box-rev">0,00 zł</div>
                    <div class="variant-stats">
                        <span>Zamówień: <strong id="lab-v-box-orders">0</strong></span>
                        <span>Sztuk: <strong id="lab-v-box-sales">0</strong></span>
                        <span>Udział: <strong id="lab-v-box-share">0%</strong></span>
                    </div>
                </div>
                <div class="variant-card">
                    <h4>Wersja Cyfrowa (On-line)</h4>
                    <div class="variant-rev" id="lab-v-digital-rev">0,00 zł</div>
                    <div class="variant-stats">
                        <span>Zamówień: <strong id="lab-v-digital-orders">0</strong></span>
                        <span>Sztuk: <strong id="lab-v-digital-sales">0</strong></span>
                        <span>Udział: <strong id="lab-v-digital-share">0%</strong></span>
                    </div>
                </div>
                <div class="variant-card">
                    <h4>Pakiet (Pudełko + On-line)</h4>
                    <div class="variant-rev" id="lab-v-bundle-rev">0,00 zł</div>
                    <div class="variant-stats">
                        <span>Zamówień: <strong id="lab-v-bundle-orders">0</strong></span>
                        <span>Sztuk: <strong id="lab-v-bundle-sales">0</strong></span>
                        <span>Udział: <strong id="lab-v-bundle-share">0%</strong></span>
                    </div>
                </div>
            </div>

            <!-- CHARTS -->
            <div class="charts-grid-row1">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Dzienna Dynamika Przychodu (zł)</h3>
                    <div style="height: 310px;"><canvas id="labChartDaily"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Przychód Skumulowany (Suma narastająco)</h3>
                    <div style="height: 310px;"><canvas id="labChartCumulative"></canvas></div>
                </div>
            </div>

            <div class="charts-grid-row2">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż Miesięczna (zł)</h3>
                    <div style="height: 310px;"><canvas id="labChartMonthly"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż Tygodniowa (zł)</h3>
                    <div style="height: 310px;"><canvas id="labChartWeekly"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż wg Dni Tygodnia</h3>
                    <div style="height: 310px;"><canvas id="labChartDayOfWeek"></canvas></div>
                </div>
            </div>

            <div class="insights">
                <h2>📊 Rekomendacje & Wnioski Analityczne: Labirynt Rozmów</h2>
                <ul>
                    <li><strong>Wersja Fizyczna dominuje przychody:</strong> Karciana wersja fizyczna stanowi ponad 78% sprzedaży gry.</li>
                    <li><strong>Up-sell na Pakiety:</strong> Wariant łączony (Pudełko + On-line) buduje wysoki koszyk AOV.</li>
                    <li><strong>Szczyty aktywności:</strong> Największa konwersja zachodzi w środku tygodnia i w godzinach wieczornych.</li>
                </ul>
            </div>
        </div>


        <!-- ========================================== -->
        <!-- TAB 2: SZKOLENIA OGÓŁEM -->
        <!-- ========================================== -->
        <div class="tab-pane" id="pane-szkolenia">
            
            <div class="filter-container">
                <div class="filter-group">
                    <span class="filter-label">Wykładowca / Autor:</span>
                    <select id="szkAuthorSelect" class="filter-input" style="cursor: pointer; font-weight: 600; padding-right: 1.5rem;">
                        <option value="all">Wszyscy Wykładowcy</option>
                        <option value="dr Ewa Woydyłło-Osiatyńska">dr Ewa Woydyłło-Osiatyńska</option>
                        <option value="Prof. Bogdan de Barbaro">Prof. Bogdan de Barbaro</option>
                        <option value="Prof. Jerzy Bralczyk">Prof. Jerzy Bralczyk</option>
                        <option value="Maciej Kołodziejczyk">Maciej Kołodziejczyk</option>
                        <option value="Jacek Walkiewicz">Jacek Walkiewicz</option>
                        <option value="Piotr Matejuk">Piotr Matejuk</option>
                        <option value="Ewa Dobiała">Ewa Dobiała</option>
                        <option value="dr Krzysztof Piotrowski">dr Krzysztof Piotrowski</option>
                        <option value="dr Maciej Dębski">dr Maciej Dębski</option>
                        <option value="Piotr Mosak">Piotr Mosak</option>
                        <option value="Adrianna Lipiarz">Adrianna Lipiarz</option>
                        <option value="Pakiety Wieloautorskie">Pakiety Wieloautorskie</option>
                        <option value="Slowdating">Slowdating (Wydarzenia)</option>
                    </select>
                </div>
                
                <div class="filter-group">
                    <span class="filter-label">Zakres dat:</span>
                    <input type="date" id="szkStartDate" class="filter-input">
                    <span class="filter-label" style="font-weight: normal; color: #555;">do</span>
                    <input type="date" id="szkEndDate" class="filter-input">
                </div>
                <div class="btn-group">
                    <button class="filter-btn" id="szk-btn-14d">Ostatnie 14 dni</button>
                    <button class="filter-btn" id="szk-btn-30d">Ostatnie 30 dni</button>
                    <button class="filter-btn" id="szk-btn-90d">Ostatnie 90 dni</button>
                    <button class="filter-btn active" id="szk-btn-all">Cały okres</button>
                </div>
            </div>
            
            <h2 class="section-title">🎓 Podsumowanie Wyników: Szkolenia On-line i Pakiety</h2>
            <div class="summary-grid">
                <div class="card">
                    <h3>Łączny Przychód ze Szkoleń</h3>
                    <div class="value accent data-value" id="szk-total-revenue">0,00 zł</div>
                    <p>Uwzględnia wszystkie kursy</p>
                </div>
                <div class="card">
                    <h3>Zamówienia Szkoleń</h3>
                    <div class="value data-value" id="szk-total-orders">0</div>
                    <p>Potwierdzone zakupy</p>
                </div>
                <div class="card">
                    <h3>Sprzedane Kursy (Sztuk)</h3>
                    <div class="value data-value" id="szk-total-sales">0</div>
                    <p>Suma wolumenu</p>
                </div>
                <div class="card">
                    <h3>Średni Koszyk (AOV)</h3>
                    <div class="value data-value" id="szk-aov">0,00 zł</div>
                    <p>Średnia transakcja</p>
                </div>
                <div class="card">
                    <h3>Bestsellerowe Szkolenie</h3>
                    <div class="value data-value" style="font-size: 1.15rem; line-height: 1.3;" id="szk-top-course">-</div>
                    <p id="szk-top-course-rev">0,00 zł</p>
                </div>
                <div class="card">
                    <h3>Top Wykładowca</h3>
                    <div class="value data-value gold" style="font-size: 1.25rem;" id="szk-top-author">-</div>
                    <p id="szk-top-author-rev">0,00 zł</p>
                </div>
            </div>

            <!-- TOP AUTHORS CARDS -->
            <h2 class="section-title">🏆 Najpopularniejsi Wykładowcy EUPHIRE</h2>
            <div class="variant-grid" id="szk-authors-cards-container"></div>

            <!-- CHARTS -->
            <div class="charts-grid-row1">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Dzienna Dynamika Przychodu ze Szkoleń (zł)</h3>
                    <div style="height: 310px;"><canvas id="szkChartDaily"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Przychód Skumulowany ze Szkoleń (zł)</h3>
                    <div style="height: 310px;"><canvas id="szkChartCumulative"></canvas></div>
                </div>
            </div>

            <div class="charts-grid-row2">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Udział Wykładowców w Przychodzie (%)</h3>
                    <div style="height: 310px;"><canvas id="szkChartAuthors"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Top 10 Najpopularniejszych Szkoleń (zł)</h3>
                    <div style="height: 310px;"><canvas id="szkChartTopCourses"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż Szkoleń wg Dni Tygodnia</h3>
                    <div style="height: 310px;"><canvas id="szkChartDayOfWeek"></canvas></div>
                </div>
            </div>

            <!-- TABLE OF COURSES -->
            <h2 class="section-title">📋 Pełne Zestawienie Wyników Szkoleń i Pakietów</h2>
            <div class="table-container">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Nazwa Szkolenia / Pakietu</th>
                            <th>Wykładowca / Autor</th>
                            <th>Zamówień</th>
                            <th>Sztuk</th>
                            <th>Przychód Brutto</th>
                            <th>Średnia Cena / Szt.</th>
                            <th>Udział w Sprzedaży</th>
                        </tr>
                    </thead>
                    <tbody id="szk-courses-table-body"></tbody>
                </table>
            </div>

            <div class="insights">
                <h2>📈 Strategiczne Wnioski dla Sprzedaży Szkoleń</h2>
                <ul>
                    <li><strong>Woydyłło-Osiatyńska liderem przychodów:</strong> Kurs o depresji generuje ponad 40% całkowitej sprzedaży szkoleń.</li>
                    <li><strong>Duety Profesorów:</strong> Pakiety łączące psychoterapię (de Barbaro) i komunikację (Bralczyk) osiągają świetny wskaźnik AOV.</li>
                </ul>
            </div>
        </div>


        <!-- ========================================== -->
        <!-- TAB 3: SZKOLENIE PIOTR MOSAK -->
        <!-- ========================================== -->
        <div class="tab-pane" id="pane-mosak">
            
            <div class="hero-banner">
                <span class="hero-badge">NOWOŚĆ W OFERCIE • PREMIERA LIPIEC 2026</span>
                <h2>Piotr Mosak – „Fundamenty dobrego związku”</h2>
                <p>Szkolenie on-line dedykowane relacjom, budowaniu więzi oraz rozwiązywaniu kryzysów w związku. Analiza sprzedażowa od dnia premiery (23 lipca 2026 r.).</p>
            </div>

            <div class="filter-container">
                <div class="filter-group">
                    <span class="filter-label">Zakres dat:</span>
                    <input type="date" id="mosStartDate" class="filter-input">
                    <span class="filter-label" style="font-weight: normal; color: #555;">do</span>
                    <input type="date" id="mosEndDate" class="filter-input">
                </div>
                <div class="btn-group">
                    <button class="filter-btn" id="mos-btn-14d">Ostatnie 14 dni</button>
                    <button class="filter-btn" id="mos-btn-30d">Ostatnie 30 dni</button>
                    <button class="filter-btn" id="mos-btn-90d">Ostatnie 90 dni</button>
                    <button class="filter-btn active" id="mos-btn-all">Cały okres</button>
                </div>
            </div>
            
            <h2 class="section-title">👤 Wyniki Sprzedaży Kursu Piotra Mosaka</h2>
            <div class="summary-grid">
                <div class="card">
                    <h3>Przychód z Kursu</h3>
                    <div class="value accent data-value" id="mos-total-revenue">0,00 zł</div>
                    <p>Faktyczne wpłaty klientów</p>
                </div>
                <div class="card">
                    <h3>Liczba Zamówień</h3>
                    <div class="value data-value" id="mos-total-orders">0</div>
                    <p>Liczba kupujących</p>
                </div>
                <div class="card">
                    <h3>Sprzedane Sztuki</h3>
                    <div class="value data-value" id="mos-total-sales">0</div>
                    <p>Dostępy online</p>
                </div>
                <div class="card">
                    <h3>Średnia Cena (AOV)</h3>
                    <div class="value data-value" id="mos-aov">0,00 zł</div>
                    <p>Średnio na zamówienie</p>
                </div>
                <div class="card">
                    <h3>Średni Przychód Dzienny</h3>
                    <div class="value data-value gold" id="mos-daily-avg">0,00 zł</div>
                    <p>Od premiery (23.07)</p>
                </div>
                <div class="card">
                    <h3>Najlepszy Dzień</h3>
                    <div class="value data-value" style="font-size: 1.3rem;" id="mos-best-day">-</div>
                    <p id="mos-best-day-rev">0,00 zł</p>
                </div>
            </div>

            <!-- CHARTS -->
            <div class="charts-grid-row1">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Dzienna Sprzedaż i Przychód Skumulowany Od Premiery</h3>
                    <div style="height: 310px;"><canvas id="mosChartDaily"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż Kursu wg Dni Tygodnia</h3>
                    <div style="height: 310px;"><canvas id="mosChartDayOfWeek"></canvas></div>
                </div>
            </div>

            <!-- TRANSACTIONS TABLE -->
            <h2 class="section-title">📑 Historia Wszystkich Transakcji Kursu Piotra Mosaka</h2>
            <div class="table-container">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Data Zamówienia</th>
                            <th>Godzina</th>
                            <th>Nazwa Produktu</th>
                            <th>Ilość</th>
                            <th>Kwota Zapłacona</th>
                            <th>Cena Nominalna</th>
                            <th>Poziom Zniżki / Kodu</th>
                        </tr>
                    </thead>
                    <tbody id="mos-table-body"></tbody>
                </table>
            </div>

            <div class="insights">
                <h2>💡 Rekomendacje Analityczne i Strategia Skalowania Sprzedaży</h2>
                <ul>
                    <li><strong>Synergia z „Labiryntem Rozmów”:</strong> Rekomendujemy cross-sell z grą dla par na Thank You Page oraz w sekwencji e-mail po zakupie.</li>
                    <li><strong>Stworzenie Pakietu „Dobry Związek”:</strong> Dodanie w sklepie pakietu łączonego (*Gra Labirynt Rozmów + Kurs Piotra Mosaka*) podniesie AOV o ponad 60%.</li>
                </ul>
            </div>
        </div>


        <!-- ========================================== -->
        <!-- TAB 4: PROF. BOGDAN DE BARBARO - DOBRY TERAPEUTA -->
        <!-- ========================================== -->
        <div class="tab-pane" id="pane-debarbaro">
            
            <div class="hero-banner">
                <span class="hero-badge">BESTSELLER PSYCHOTERAPII</span>
                <h2>Prof. Bogdan de Barbaro – „Jak być dobrym psychoterapeutą dla samego siebie?”</h2>
                <p>Jeden z największych bestsellerów platformy EUPHIRE. Przystępny, mądry i głęboki kurs uczy autorefleksji, przepracowywania konfliktów wewnętrznych oraz higieny psychicznej w oparciu o dorobek profesorski.</p>
            </div>

            <div class="filter-container">
                <div class="filter-group">
                    <span class="filter-label">Zakres dat:</span>
                    <input type="date" id="barStartDate" class="filter-input">
                    <span class="filter-label" style="font-weight: normal; color: #555;">do</span>
                    <input type="date" id="barEndDate" class="filter-input">
                </div>
                <div class="btn-group">
                    <button class="filter-btn" id="bar-btn-14d">Ostatnie 14 dni</button>
                    <button class="filter-btn" id="bar-btn-30d">Ostatnie 30 dni</button>
                    <button class="filter-btn" id="bar-btn-90d">Ostatnie 90 dni</button>
                    <button class="filter-btn active" id="bar-btn-all">Cały okres</button>
                </div>
            </div>
            
            <h2 class="section-title">👨‍⚕️ Wyniki Sprzedaży: Prof. de Barbaro – „Dobry Terapeuta”</h2>
            <div class="summary-grid">
                <div class="card">
                    <h3>Całkowity Przychód</h3>
                    <div class="value accent data-value" id="bar-total-revenue">0,00 zł</div>
                    <p>Przychód brutto ze szkolenia</p>
                </div>
                <div class="card">
                    <h3>Liczba Zamówień</h3>
                    <div class="value data-value" id="bar-total-orders">0</div>
                    <p>Sprzedanych dostępów</p>
                </div>
                <div class="card">
                    <h3>Sprzedane Sztuki</h3>
                    <div class="value data-value" id="bar-total-sales">0</div>
                    <p>Łączny wolumen</p>
                </div>
                <div class="card">
                    <h3>Średni Koszyk (AOV)</h3>
                    <div class="value data-value" id="bar-aov">0,00 zł</div>
                    <p>Średnia cena transakcyjna</p>
                </div>
                <div class="card">
                    <h3>Najlepszy Dzień</h3>
                    <div class="value data-value" style="font-size: 1.3rem;" id="bar-best-day">-</div>
                    <p id="bar-best-day-rev">0,00 zł</p>
                </div>
                <div class="card">
                    <h3>Udział w Szkoleniach</h3>
                    <div class="value data-value gold" id="bar-share-perc">0%</div>
                    <p>Wielkość udziału w przychodach</p>
                </div>
            </div>

            <!-- CHARTS -->
            <div class="charts-grid-row1">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Dzienna Dynamika Przychodu (zł)</h3>
                    <div style="height: 310px;"><canvas id="barChartDaily"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Przychód Skumulowany (zł)</h3>
                    <div style="height: 310px;"><canvas id="barChartCumulative"></canvas></div>
                </div>
            </div>

            <div class="charts-grid-row2">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż Miesięczna (zł)</h3>
                    <div style="height: 310px;"><canvas id="barChartMonthly"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Struktura Poziomów Cenowych (zł)</h3>
                    <div style="height: 310px;"><canvas id="barChartPriceDist"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż wg Dni Tygodnia</h3>
                    <div style="height: 310px;"><canvas id="barChartDayOfWeek"></canvas></div>
                </div>
            </div>

            <div class="insights">
                <h2>📊 Strategiczne Wnioski i Analityka: Prof. de Barbaro</h2>
                <ul>
                    <li><strong>Produkt typu Evergreen:</strong> Szkolenie generuje stabilny przychód w skali całego roku (ponad 71 tys. zł przychodów brutto).</li>
                    <li><strong>Świetny próg cenowy:</strong> Najwyższa konwersja zachodzi przy progu 199 zł oraz promocyjnym 149 zł.</li>
                    <li><strong>Budowa Kompendiów:</strong> Łączenie prof. de Barbaro w pakiety z dr Woydyłło-Osiatyńską i prof. Bralczykiem przynosi najwyższy średni AOV.</li>
                </ul>
            </div>
        </div>


        <!-- ========================================== -->
        <!-- TAB 5: JACEK WALKIEWICZ - PEŁNA MOC PRZEMAWIANIA -->
        <!-- ========================================== -->
        <div class="tab-pane" id="pane-walkiewicz">
            
            <div class="hero-banner">
                <span class="hero-badge">REKORDOWY AOV • PREMIUM HIGH-TICKET</span>
                <h2>Jacek Walkiewicz – „Pełna MOC Przemawiania”</h2>
                <p>Flagowe szkolenie autorskie kultowego mówcy motywacyjnego Jacka Walkiewicza. Produkt charakteryzuje się najwyższym średnim koszykiem zakupowym (AOV > 308 zł) w całym sklepie EUPHIRE.</p>
            </div>

            <div class="filter-container">
                <div class="filter-group">
                    <span class="filter-label">Zakres dat:</span>
                    <input type="date" id="walStartDate" class="filter-input">
                    <span class="filter-label" style="font-weight: normal; color: #555;">do</span>
                    <input type="date" id="walEndDate" class="filter-input">
                </div>
                <div class="btn-group">
                    <button class="filter-btn" id="wal-btn-14d">Ostatnie 14 dni</button>
                    <button class="filter-btn" id="wal-btn-30d">Ostatnie 30 dni</button>
                    <button class="filter-btn" id="wal-btn-90d">Ostatnie 90 dni</button>
                    <button class="filter-btn active" id="wal-btn-all">Cały okres</button>
                </div>
            </div>
            
            <h2 class="section-title">🎤 Wyniki Sprzedaży: Jacek Walkiewicz – „Pełna MOC”</h2>
            <div class="summary-grid">
                <div class="card">
                    <h3>Całkowity Przychód</h3>
                    <div class="value accent data-value" id="wal-total-revenue">0,00 zł</div>
                    <p>Przychód brutto ze szkolenia</p>
                </div>
                <div class="card">
                    <h3>Liczba Zamówień</h3>
                    <div class="value data-value" id="wal-total-orders">0</div>
                    <p>Transakcje zakupu</p>
                </div>
                <div class="card">
                    <h3>Sprzedane Sztuki</h3>
                    <div class="value data-value" id="wal-total-sales">0</div>
                    <p>Dostępy do kursu</p>
                </div>
                <div class="card">
                    <h3>Średni Koszyk (AOV)</h3>
                    <div class="value data-value gold" id="wal-aov">0,00 zł</div>
                    <p>Najwyższy AOV na platformie!</p>
                </div>
                <div class="card">
                    <h3>Najlepszy Dzień</h3>
                    <div class="value data-value" style="font-size: 1.3rem;" id="wal-best-day">-</div>
                    <p id="wal-best-day-rev">0,00 zł</p>
                </div>
                <div class="card">
                    <h3>Procent Cen Premium</h3>
                    <div class="value data-value" id="wal-premium-perc">0%</div>
                    <p>Zamówienia powyżej 299 zł</p>
                </div>
            </div>

            <!-- CHARTS -->
            <div class="charts-grid-row1">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Dzienna Dynamika Przychodu (zł)</h3>
                    <div style="height: 310px;"><canvas id="walChartDaily"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Przychód Skumulowany (zł)</h3>
                    <div style="height: 310px;"><canvas id="walChartCumulative"></canvas></div>
                </div>
            </div>

            <div class="charts-grid-row2">
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż Miesięczna (zł)</h3>
                    <div style="height: 310px;"><canvas id="walChartMonthly"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Struktura Poziomów Cenowych (zł)</h3>
                    <div style="height: 310px;"><canvas id="walChartPriceDist"></canvas></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 1rem; font-size: 1.05rem; color: #004D54;">Sprzedaż wg Dni Tygodnia</h3>
                    <div style="height: 310px;"><canvas id="walChartDayOfWeek"></canvas></div>
                </div>
            </div>

            <div class="insights">
                <h2>💡 Rekomendacje i Strategia Premium: Jacek Walkiewicz</h2>
                <ul>
                    <li><strong>Pozycja High-Ticket:</strong> Klienci bez oporów kupują kurs w cenie nominalnej 399 zł lub promocyjnej 299,25 zł.</li>
                    <li><strong>Rekomendowany Pakiet „Mistrzowie Przemawiania”:</strong> Połączenie w pakiecie szkolenia Jacka Walkiewicza z kursami prof. Jerzego Bralczyka (*„Jak dobrze mówić?”*) pozwoli wygenerować pakiet z koszykiem > 499 zł.</li>
                </ul>
            </div>
        </div>

    </div>

    <!-- DYNAMIC JAVASCRIPT LOGIC -->
    <script>
        const rawTransactions = {{RAW_TRANSACTIONS_JSON}};
        
        const formatPLN = (val) => val.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
        const formatMonth = (monthStr) => {
            const [y, m] = monthStr.split('-');
            const months = ['Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec', 'Lipiec', 'Sierpień', 'Wrzesień', 'Październik', 'Listopad', 'Grudzień'];
            return months[parseInt(m, 10) - 1] + ' ' + y;
        };
        const getDayOfWeekIndex = (dateStr) => {
            const d = new Date(dateStr + 'T00:00:00');
            let day = d.getDay();
            return day === 0 ? 6 : day - 1;
        };
        const getWeekNumberStr = (dateStr) => {
            const d = new Date(dateStr + 'T00:00:00');
            const dYear = d.getFullYear();
            const firstJan = new Date(dYear, 0, 1);
            const days = Math.floor((d - firstJan) / (24 * 60 * 60 * 1000));
            const week = Math.ceil((days + firstJan.getDay() + 1) / 7);
            return `${dYear}-W${week < 10 ? '0' + week : week}`;
        };

        const colors = {
            primary: '#004D54',
            primaryLight: 'rgba(0, 77, 84, 0.15)',
            secondary: '#002E32',
            accent: '#FCAE2F',
            accentLight: 'rgba(252, 174, 47, 0.2)',
            mist: '#B2CACC',
            text: '#1D2A2B',
            palette: ['#004D54', '#FCAE2F', '#007A87', '#D97706', '#002E32', '#B2CACC', '#0284C7', '#7C3AED', '#EC4899', '#10B981']
        };

        Chart.defaults.font.family = "'Plus Jakarta Sans', sans-serif";
        Chart.defaults.color = '#2F3E40';

        // ----------------------------------------------------
        // INITIALIZE LABIRYNT CHARTS
        // ----------------------------------------------------
        const ctxLabDaily = document.getElementById('labChartDaily').getContext('2d');
        const labChartDaily = new Chart(ctxLabDaily, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], borderColor: colors.primary, backgroundColor: colors.primaryLight, fill: true, tension: 0.3, pointRadius: 2 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxLabCumulative = document.getElementById('labChartCumulative').getContext('2d');
        const labChartCumulative = new Chart(ctxLabCumulative, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Przychód Skumulowany (zł)', data: [], borderColor: colors.accent, backgroundColor: colors.accentLight, fill: true, tension: 0.3, pointRadius: 0, borderWidth: 3 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxLabMonthly = document.getElementById('labChartMonthly').getContext('2d');
        const labChartMonthly = new Chart(ctxLabMonthly, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], backgroundColor: colors.primary, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxLabWeekly = document.getElementById('labChartWeekly').getContext('2d');
        const labChartWeekly = new Chart(ctxLabWeekly, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], backgroundColor: colors.accent, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxLabDayOfWeek = document.getElementById('labChartDayOfWeek').getContext('2d');
        const labChartDayOfWeek = new Chart(ctxLabDayOfWeek, {
            type: 'bar',
            data: { labels: ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela'], datasets: [{ label: 'Przychód (zł)', data: [0,0,0,0,0,0,0], backgroundColor: colors.primary, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        // ----------------------------------------------------
        // INITIALIZE SZKOLENIA CHARTS
        // ----------------------------------------------------
        const ctxSzkDaily = document.getElementById('szkChartDaily').getContext('2d');
        const szkChartDaily = new Chart(ctxSzkDaily, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], borderColor: colors.primary, backgroundColor: colors.primaryLight, fill: true, tension: 0.3, pointRadius: 2 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxSzkCumulative = document.getElementById('szkChartCumulative').getContext('2d');
        const szkChartCumulative = new Chart(ctxSzkCumulative, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Skumulowany (zł)', data: [], borderColor: colors.accent, backgroundColor: colors.accentLight, fill: true, tension: 0.3, pointRadius: 0, borderWidth: 3 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxSzkAuthors = document.getElementById('szkChartAuthors').getContext('2d');
        const szkChartAuthors = new Chart(ctxSzkAuthors, {
            type: 'doughnut',
            data: { labels: [], datasets: [{ data: [], backgroundColor: colors.palette }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } } } }
        });

        const ctxSzkTopCourses = document.getElementById('szkChartTopCourses').getContext('2d');
        const szkChartTopCourses = new Chart(ctxSzkTopCourses, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], backgroundColor: colors.primary, borderRadius: 6 }] },
            options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxSzkDayOfWeek = document.getElementById('szkChartDayOfWeek').getContext('2d');
        const szkChartDayOfWeek = new Chart(ctxSzkDayOfWeek, {
            type: 'bar',
            data: { labels: ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela'], datasets: [{ label: 'Przychód (zł)', data: [0,0,0,0,0,0,0], backgroundColor: colors.accent, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        // ----------------------------------------------------
        // INITIALIZE MOSAK CHARTS
        // ----------------------------------------------------
        const ctxMosDaily = document.getElementById('mosChartDaily').getContext('2d');
        const mosChartDaily = new Chart(ctxMosDaily, {
            type: 'bar',
            data: {
                labels: [],
                datasets: [
                    { type: 'line', label: 'Przychód Skumulowany (zł)', data: [], borderColor: colors.accent, backgroundColor: colors.accent, borderWidth: 3, tension: 0.2, yAxisID: 'y1' },
                    { type: 'bar', label: 'Przychód Dzienny (zł)', data: [], backgroundColor: colors.primary, borderRadius: 6, yAxisID: 'y' }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: { type: 'linear', display: true, position: 'left', title: { display: true, text: 'Dzienna sprzedaż (zł)' } },
                    y1: { type: 'linear', display: true, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Skumulowany (zł)' } }
                }
            }
        });

        const ctxMosDayOfWeek = document.getElementById('mosChartDayOfWeek').getContext('2d');
        const mosChartDayOfWeek = new Chart(ctxMosDayOfWeek, {
            type: 'bar',
            data: { labels: ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela'], datasets: [{ label: 'Liczba Zamówień', data: [0,0,0,0,0,0,0], backgroundColor: colors.accent, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        // ----------------------------------------------------
        // INITIALIZE DE BARBARO TERAPEUTA CHARTS
        // ----------------------------------------------------
        const ctxBarDaily = document.getElementById('barChartDaily').getContext('2d');
        const barChartDaily = new Chart(ctxBarDaily, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], borderColor: colors.primary, backgroundColor: colors.primaryLight, fill: true, tension: 0.3, pointRadius: 2 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxBarCumulative = document.getElementById('barChartCumulative').getContext('2d');
        const barChartCumulative = new Chart(ctxBarCumulative, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Skumulowany (zł)', data: [], borderColor: colors.accent, backgroundColor: colors.accentLight, fill: true, tension: 0.3, pointRadius: 0, borderWidth: 3 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxBarMonthly = document.getElementById('barChartMonthly').getContext('2d');
        const barChartMonthly = new Chart(ctxBarMonthly, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], backgroundColor: colors.primary, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxBarPriceDist = document.getElementById('barChartPriceDist').getContext('2d');
        const barChartPriceDist = new Chart(ctxBarPriceDist, {
            type: 'doughnut',
            data: { labels: ['Cena standardowa (199 zł)', 'Promocja (149 zł)', 'Kod rabatowy (179,10 zł)', 'Rabaty -50% (99,50 zł)', 'Inne'], datasets: [{ data: [0,0,0,0,0], backgroundColor: colors.palette }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12 } } } }
        });

        const ctxBarDayOfWeek = document.getElementById('barChartDayOfWeek').getContext('2d');
        const barChartDayOfWeek = new Chart(ctxBarDayOfWeek, {
            type: 'bar',
            data: { labels: ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela'], datasets: [{ label: 'Przychód (zł)', data: [0,0,0,0,0,0,0], backgroundColor: colors.accent, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        // ----------------------------------------------------
        // INITIALIZE WALKIEWICZ CHARTS
        // ----------------------------------------------------
        const ctxWalDaily = document.getElementById('walChartDaily').getContext('2d');
        const walChartDaily = new Chart(ctxWalDaily, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], borderColor: colors.primary, backgroundColor: colors.primaryLight, fill: true, tension: 0.3, pointRadius: 2 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxWalCumulative = document.getElementById('walChartCumulative').getContext('2d');
        const walChartCumulative = new Chart(ctxWalCumulative, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Skumulowany (zł)', data: [], borderColor: colors.accent, backgroundColor: colors.accentLight, fill: true, tension: 0.3, pointRadius: 0, borderWidth: 3 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxWalMonthly = document.getElementById('walChartMonthly').getContext('2d');
        const walChartMonthly = new Chart(ctxWalMonthly, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'Przychód (zł)', data: [], backgroundColor: colors.primary, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });

        const ctxWalPriceDist = document.getElementById('walChartPriceDist').getContext('2d');
        const walChartPriceDist = new Chart(ctxWalPriceDist, {
            type: 'doughnut',
            data: { labels: ['Rabat 25% (299,25 zł)', 'Cena Pełna (399 zł)', 'Cena Pro (254,25 zł)', 'Rabat VIP (339 zł)', 'Inne'], datasets: [{ data: [0,0,0,0,0], backgroundColor: colors.palette }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12 } } } }
        });

        const ctxWalDayOfWeek = document.getElementById('walChartDayOfWeek').getContext('2d');
        const walChartDayOfWeek = new Chart(ctxWalDayOfWeek, {
            type: 'bar',
            data: { labels: ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela'], datasets: [{ label: 'Przychód (zł)', data: [0,0,0,0,0,0,0], backgroundColor: colors.accent, borderRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });


        // ----------------------------------------------------
        // UPDATE LOGIC - TAB 1: LABIRYNT
        // ----------------------------------------------------
        function updateLabiryntDashboard() {
            const variantFilter = document.getElementById('labVariantSelect').value;
            const startDate = document.getElementById('labStartDate').value;
            const endDate = document.getElementById('labEndDate').value;

            const filtered = rawTransactions.filter(t => {
                if (!t.isL) return false;
                if (variantFilter !== 'all' && t.vL !== variantFilter) return false;
                if (startDate && t.d < startDate) return false;
                if (endDate && t.d > endDate) return false;
                return true;
            });

            let totalRev = 0, totalOrders = filtered.length, totalSales = 0;
            let vBoxRev = 0, vBoxOrders = 0, vBoxSales = 0;
            let vDigRev = 0, vDigOrders = 0, vDigSales = 0;
            let vBunRev = 0, vBunOrders = 0, vBunSales = 0;

            const dailyMap = {}, monthlyMap = {}, weeklyMap = {};
            const dayOfWeekMap = [0, 0, 0, 0, 0, 0, 0];

            filtered.forEach(t => {
                totalRev += t.k;
                totalSales += t.s;

                if (t.vL.includes('Fizyczna')) { vBoxRev += t.k; vBoxOrders++; vBoxSales += t.s; }
                else if (t.vL.includes('Cyfrową')) { vDigRev += t.k; vDigOrders++; vDigSales += t.s; }
                else if (t.vL.includes('Pakiet')) { vBunRev += t.k; vBunOrders++; vBunSales += t.s; }

                dailyMap[t.d] = (dailyMap[t.d] || 0) + t.k;
                const monthKey = t.d.substring(0, 7);
                monthlyMap[monthKey] = (monthlyMap[monthKey] || 0) + t.k;

                const weekKey = getWeekNumberStr(t.d);
                weeklyMap[weekKey] = (weeklyMap[weekKey] || 0) + t.k;

                const dow = getDayOfWeekIndex(t.d);
                dayOfWeekMap[dow] += t.k;
            });

            document.getElementById('lab-total-revenue').innerText = formatPLN(totalRev);
            document.getElementById('lab-total-orders').innerText = totalOrders.toLocaleString('pl-PL');
            document.getElementById('lab-total-sales').innerText = totalSales.toLocaleString('pl-PL');
            document.getElementById('lab-aov').innerText = totalOrders > 0 ? formatPLN(totalRev / totalOrders) : '0,00 zł';

            document.getElementById('lab-v-box-rev').innerText = formatPLN(vBoxRev);
            document.getElementById('lab-v-box-orders').innerText = vBoxOrders;
            document.getElementById('lab-v-box-sales').innerText = vBoxSales;
            document.getElementById('lab-v-box-share').innerText = totalRev > 0 ? ((vBoxRev / totalRev) * 100).toFixed(1) + '%' : '0%';

            document.getElementById('lab-v-digital-rev').innerText = formatPLN(vDigRev);
            document.getElementById('lab-v-digital-orders').innerText = vDigOrders;
            document.getElementById('lab-v-digital-sales').innerText = vDigSales;
            document.getElementById('lab-v-digital-share').innerText = totalRev > 0 ? ((vDigRev / totalRev) * 100).toFixed(1) + '%' : '0%';

            document.getElementById('lab-v-bundle-rev').innerText = formatPLN(vBunRev);
            document.getElementById('lab-v-bundle-orders').innerText = vBunOrders;
            document.getElementById('lab-v-bundle-sales').innerText = vBunSales;
            document.getElementById('lab-v-bundle-share').innerText = totalRev > 0 ? ((vBunRev / totalRev) * 100).toFixed(1) + '%' : '0%';

            let bestDay = '-', bestDayRev = 0;
            Object.keys(dailyMap).forEach(d => { if (dailyMap[d] > bestDayRev) { bestDay = d; bestDayRev = dailyMap[d]; } });
            let bestWeek = '-', bestWeekRev = 0;
            Object.keys(weeklyMap).forEach(w => { if (weeklyMap[w] > bestWeekRev) { bestWeek = w; bestWeekRev = weeklyMap[w]; } });

            document.getElementById('lab-best-day-badge').innerText = bestDay;
            document.getElementById('lab-best-day-rev').innerText = formatPLN(bestDayRev);
            document.getElementById('lab-best-week-badge').innerText = bestWeek;
            document.getElementById('lab-best-week-rev').innerText = formatPLN(bestWeekRev);

            const sortedDays = Object.keys(dailyMap).sort();
            labChartDaily.data.labels = sortedDays;
            labChartDaily.data.datasets[0].data = sortedDays.map(k => dailyMap[k]);
            labChartDaily.update();

            let cumSum = 0;
            const cumData = sortedDays.map(d => { cumSum += dailyMap[d]; return cumSum; });
            labChartCumulative.data.labels = sortedDays;
            labChartCumulative.data.datasets[0].data = cumData;
            labChartCumulative.update();

            const sortedMonths = Object.keys(monthlyMap).sort();
            labChartMonthly.data.labels = sortedMonths.map(formatMonth);
            labChartMonthly.data.datasets[0].data = sortedMonths.map(k => monthlyMap[k]);
            labChartMonthly.update();

            const sortedWeeks = Object.keys(weeklyMap).sort();
            labChartWeekly.data.labels = sortedWeeks;
            labChartWeekly.data.datasets[0].data = sortedWeeks.map(k => weeklyMap[k]);
            labChartWeekly.update();

            labChartDayOfWeek.data.datasets[0].data = dayOfWeekMap;
            labChartDayOfWeek.update();
        }

        // ----------------------------------------------------
        // UPDATE LOGIC - TAB 2: SZKOLENIA OGÓŁEM
        // ----------------------------------------------------
        function updateSzkoleniaDashboard() {
            const authorFilter = document.getElementById('szkAuthorSelect').value;
            const startDate = document.getElementById('szkStartDate').value;
            const endDate = document.getElementById('szkEndDate').value;

            const filtered = rawTransactions.filter(t => {
                if (t.isL) return false;
                if (authorFilter !== 'all' && t.a !== authorFilter) return false;
                if (startDate && t.d < startDate) return false;
                if (endDate && t.d > endDate) return false;
                return true;
            });

            let totalRev = 0, totalOrders = filtered.length, totalSales = 0;
            const dailyMap = {}, courseMap = {}, authorMap = {};
            const dayOfWeekMap = [0, 0, 0, 0, 0, 0, 0];

            filtered.forEach(t => {
                totalRev += t.k;
                totalSales += t.s;

                dailyMap[t.d] = (dailyMap[t.d] || 0) + t.k;

                if (!courseMap[t.n]) {
                    courseMap[t.n] = { name: t.n, author: t.a, rev: 0, orders: 0, sales: 0 };
                }
                courseMap[t.n].rev += t.k;
                courseMap[t.n].orders++;
                courseMap[t.n].sales += t.s;

                if (!authorMap[t.a]) {
                    authorMap[t.a] = { name: t.a, rev: 0, orders: 0, sales: 0 };
                }
                authorMap[t.a].rev += t.k;
                authorMap[t.a].orders++;
                authorMap[t.a].sales += t.s;

                const dow = getDayOfWeekIndex(t.d);
                dayOfWeekMap[dow] += t.k;
            });

            document.getElementById('szk-total-revenue').innerText = formatPLN(totalRev);
            document.getElementById('szk-total-orders').innerText = totalOrders.toLocaleString('pl-PL');
            document.getElementById('szk-total-sales').innerText = totalSales.toLocaleString('pl-PL');
            document.getElementById('szk-aov').innerText = totalOrders > 0 ? formatPLN(totalRev / totalOrders) : '0,00 zł';

            const sortedCourses = Object.values(courseMap).sort((a, b) => b.rev - a.rev);
            const sortedAuthors = Object.values(authorMap).sort((a, b) => b.rev - a.rev);

            if (sortedCourses.length > 0) {
                document.getElementById('szk-top-course').innerText = sortedCourses[0].name;
                document.getElementById('szk-top-course-rev').innerText = formatPLN(sortedCourses[0].rev);
            } else {
                document.getElementById('szk-top-course').innerText = '-';
                document.getElementById('szk-top-course-rev').innerText = '0,00 zł';
            }

            if (sortedAuthors.length > 0) {
                document.getElementById('szk-top-author').innerText = sortedAuthors[0].name;
                document.getElementById('szk-top-author-rev').innerText = formatPLN(sortedAuthors[0].rev);
            } else {
                document.getElementById('szk-top-author').innerText = '-';
                document.getElementById('szk-top-author-rev').innerText = '0,00 zł';
            }

            const authorsContainer = document.getElementById('szk-authors-cards-container');
            authorsContainer.innerHTML = '';
            sortedAuthors.slice(0, 4).forEach(aut => {
                const share = totalRev > 0 ? ((aut.rev / totalRev) * 100).toFixed(1) : 0;
                const card = document.createElement('div');
                card.className = 'variant-card';
                card.innerHTML = `
                    <h4>${aut.name}</h4>
                    <div class="variant-rev">${formatPLN(aut.rev)}</div>
                    <div class="variant-stats">
                        <span>Zamówień: <strong>${aut.orders}</strong></span>
                        <span>Sztuk: <strong>${aut.sales}</strong></span>
                        <span>Udział: <strong>${share}%</strong></span>
                    </div>
                `;
                authorsContainer.appendChild(card);
            });

            const sortedDays = Object.keys(dailyMap).sort();
            szkChartDaily.data.labels = sortedDays;
            szkChartDaily.data.datasets[0].data = sortedDays.map(k => dailyMap[k]);
            szkChartDaily.update();

            let cumSum = 0;
            const cumData = sortedDays.map(d => { cumSum += dailyMap[d]; return cumSum; });
            szkChartCumulative.data.labels = sortedDays;
            szkChartCumulative.data.datasets[0].data = cumData;
            szkChartCumulative.update();

            szkChartAuthors.data.labels = sortedAuthors.map(a => a.name);
            szkChartAuthors.data.datasets[0].data = sortedAuthors.map(a => a.rev);
            szkChartAuthors.update();

            const top10Courses = sortedCourses.slice(0, 10);
            szkChartTopCourses.data.labels = top10Courses.map(c => c.name.length > 35 ? c.name.substring(0, 35) + '...' : c.name);
            szkChartTopCourses.data.datasets[0].data = top10Courses.map(c => c.rev);
            szkChartTopCourses.update();

            szkChartDayOfWeek.data.datasets[0].data = dayOfWeekMap;
            szkChartDayOfWeek.update();

            const tableBody = document.getElementById('szk-courses-table-body');
            tableBody.innerHTML = '';
            sortedCourses.forEach(c => {
                const avgPrice = c.sales > 0 ? c.rev / c.sales : 0;
                const share = totalRev > 0 ? ((c.rev / totalRev) * 100).toFixed(1) : 0;
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="font-weight: 600;">${c.name}</td>
                    <td><span class="badge-tag">${c.author}</span></td>
                    <td>${c.orders}</td>
                    <td>${c.sales}</td>
                    <td style="font-weight: 700; color: #004D54;">${formatPLN(c.rev)}</td>
                    <td>${formatPLN(avgPrice)}</td>
                    <td><strong>${share}%</strong></td>
                `;
                tableBody.appendChild(tr);
            });
        }

        // ----------------------------------------------------
        // UPDATE LOGIC - TAB 3: MOSAK
        // ----------------------------------------------------
        function updateMosakDashboard() {
            const startDate = document.getElementById('mosStartDate').value;
            const endDate = document.getElementById('mosEndDate').value;

            const filtered = rawTransactions.filter(t => {
                if (!t.isM) return false;
                if (startDate && t.d < startDate) return false;
                if (endDate && t.d > endDate) return false;
                return true;
            }).sort((a, b) => (a.d + a.g).localeCompare(b.d + b.g));

            let totalRev = 0, totalOrders = filtered.length, totalSales = 0;
            const dailyMap = {};
            const dayOfWeekMap = [0, 0, 0, 0, 0, 0, 0];

            filtered.forEach(t => {
                totalRev += t.k;
                totalSales += t.s;
                dailyMap[t.d] = (dailyMap[t.d] || 0) + t.k;
                const dow = getDayOfWeekIndex(t.d);
                dayOfWeekMap[dow]++;
            });

            document.getElementById('mos-total-revenue').innerText = formatPLN(totalRev);
            document.getElementById('mos-total-orders').innerText = totalOrders;
            document.getElementById('mos-total-sales').innerText = totalSales;
            document.getElementById('mos-aov').innerText = totalOrders > 0 ? formatPLN(totalRev / totalOrders) : '0,00 zł';

            const activeDaysCount = Object.keys(dailyMap).length;
            document.getElementById('mos-daily-avg').innerText = activeDaysCount > 0 ? formatPLN(totalRev / activeDaysCount) : '0,00 zł';

            let bestDay = '-', bestDayRev = 0;
            Object.keys(dailyMap).forEach(d => {
                if (dailyMap[d] > bestDayRev) { bestDay = d; bestDayRev = dailyMap[d]; }
            });
            document.getElementById('mos-best-day').innerText = bestDay;
            document.getElementById('mos-best-day-rev').innerText = formatPLN(bestDayRev);

            const sortedDays = Object.keys(dailyMap).sort();
            let cumSum = 0;
            const cumData = sortedDays.map(d => { cumSum += dailyMap[d]; return cumSum; });

            mosChartDaily.data.labels = sortedDays;
            mosChartDaily.data.datasets[0].data = cumData;
            mosChartDaily.data.datasets[1].data = sortedDays.map(k => dailyMap[k]);
            mosChartDaily.update();

            mosChartDayOfWeek.data.datasets[0].data = dayOfWeekMap;
            mosChartDayOfWeek.update();

            const tableBody = document.getElementById('mos-table-body');
            tableBody.innerHTML = '';
            filtered.forEach(t => {
                const tr = document.createElement('tr');
                const nominalPrice = 149.00;
                let discountText = 'Cena standardowa (149 zł)';
                if (t.k < 130) discountText = 'Rabat VIP (-20%)';
                else if (t.k < 140) discountText = 'Kod promocyjny (-10%)';

                tr.innerHTML = `
                    <td style="font-weight: 600;">${t.d}</td>
                    <td>${t.g}</td>
                    <td>${t.n}</td>
                    <td>${t.s}</td>
                    <td style="font-weight: 700; color: #004D54;">${formatPLN(t.k)}</td>
                    <td>${formatPLN(nominalPrice)}</td>
                    <td><span class="badge-tag">${discountText}</span></td>
                `;
                tableBody.appendChild(tr);
            });
        }

        // ----------------------------------------------------
        // UPDATE LOGIC - TAB 4: DE BARBARO (DOBRY TERAPEUTA)
        // ----------------------------------------------------
        function updateDebarbaroDashboard() {
            const startDate = document.getElementById('barStartDate').value;
            const endDate = document.getElementById('barEndDate').value;

            const filtered = rawTransactions.filter(t => {
                if (!t.isB) return false;
                if (startDate && t.d < startDate) return false;
                if (endDate && t.d > endDate) return false;
                return true;
            });

            const allSzkoleniaRev = rawTransactions.filter(t => !t.isL).reduce((acc, curr) => acc + curr.k, 0);

            let totalRev = 0, totalOrders = filtered.length, totalSales = 0;
            const dailyMap = {}, monthlyMap = {};
            const dayOfWeekMap = [0, 0, 0, 0, 0, 0, 0];
            let price199 = 0, price149 = 0, price179 = 0, price99 = 0, priceOther = 0;

            filtered.forEach(t => {
                totalRev += t.k;
                totalSales += t.s;
                dailyMap[t.d] = (dailyMap[t.d] || 0) + t.k;
                const mKey = t.d.substring(0, 7);
                monthlyMap[mKey] = (monthlyMap[mKey] || 0) + t.k;

                const dow = getDayOfWeekIndex(t.d);
                dayOfWeekMap[dow] += t.k;

                if (Math.abs(t.k - 199.00) < 1) price199 += t.k;
                else if (Math.abs(t.k - 149.00) < 1) price149 += t.k;
                else if (Math.abs(t.k - 179.10) < 1) price179 += t.k;
                else if (Math.abs(t.k - 99.50) < 1) price99 += t.k;
                else priceOther += t.k;
            });

            document.getElementById('bar-total-revenue').innerText = formatPLN(totalRev);
            document.getElementById('bar-total-orders').innerText = totalOrders;
            document.getElementById('bar-total-sales').innerText = totalSales;
            document.getElementById('bar-aov').innerText = totalOrders > 0 ? formatPLN(totalRev / totalOrders) : '0,00 zł';
            document.getElementById('bar-share-perc').innerText = allSzkoleniaRev > 0 ? ((totalRev / allSzkoleniaRev) * 100).toFixed(1) + '%' : '0%';

            let bestDay = '-', bestDayRev = 0;
            Object.keys(dailyMap).forEach(d => {
                if (dailyMap[d] > bestDayRev) { bestDay = d; bestDayRev = dailyMap[d]; }
            });
            document.getElementById('bar-best-day').innerText = bestDay;
            document.getElementById('bar-best-day-rev').innerText = formatPLN(bestDayRev);

            const sortedDays = Object.keys(dailyMap).sort();
            barChartDaily.data.labels = sortedDays;
            barChartDaily.data.datasets[0].data = sortedDays.map(k => dailyMap[k]);
            barChartDaily.update();

            let cumSum = 0;
            const cumData = sortedDays.map(d => { cumSum += dailyMap[d]; return cumSum; });
            barChartCumulative.data.labels = sortedDays;
            barChartCumulative.data.datasets[0].data = cumData;
            barChartCumulative.update();

            const sortedMonths = Object.keys(monthlyMap).sort();
            barChartMonthly.data.labels = sortedMonths.map(formatMonth);
            barChartMonthly.data.datasets[0].data = sortedMonths.map(k => monthlyMap[k]);
            barChartMonthly.update();

            barChartPriceDist.data.datasets[0].data = [price199, price149, price179, price99, priceOther];
            barChartPriceDist.update();

            barChartDayOfWeek.data.datasets[0].data = dayOfWeekMap;
            barChartDayOfWeek.update();
        }

        // ----------------------------------------------------
        // UPDATE LOGIC - TAB 5: WALKIEWICZ
        // ----------------------------------------------------
        function updateWalkiewiczDashboard() {
            const startDate = document.getElementById('walStartDate').value;
            const endDate = document.getElementById('walEndDate').value;

            const filtered = rawTransactions.filter(t => {
                if (!t.isW) return false;
                if (startDate && t.d < startDate) return false;
                if (endDate && t.d > endDate) return false;
                return true;
            });

            let totalRev = 0, totalOrders = filtered.length, totalSales = 0;
            let premiumOrdersCount = 0;
            const dailyMap = {}, monthlyMap = {};
            const dayOfWeekMap = [0, 0, 0, 0, 0, 0, 0];

            let p299 = 0, p399 = 0, p254 = 0, p339 = 0, pOther = 0;

            filtered.forEach(t => {
                totalRev += t.k;
                totalSales += t.s;
                if (t.k >= 290) premiumOrdersCount++;

                dailyMap[t.d] = (dailyMap[t.d] || 0) + t.k;
                const mKey = t.d.substring(0, 7);
                monthlyMap[mKey] = (monthlyMap[mKey] || 0) + t.k;

                const dow = getDayOfWeekIndex(t.d);
                dayOfWeekMap[dow] += t.k;

                if (Math.abs(t.k - 299.25) < 2) p299 += t.k;
                else if (Math.abs(t.k - 399.00) < 2) p399 += t.k;
                else if (Math.abs(t.k - 254.25) < 2) p254 += t.k;
                else if (Math.abs(t.k - 339.00) < 2) p339 += t.k;
                else pOther += t.k;
            });

            document.getElementById('wal-total-revenue').innerText = formatPLN(totalRev);
            document.getElementById('wal-total-orders').innerText = totalOrders;
            document.getElementById('wal-total-sales').innerText = totalSales;
            document.getElementById('wal-aov').innerText = totalOrders > 0 ? formatPLN(totalRev / totalOrders) : '0,00 zł';
            document.getElementById('wal-premium-perc').innerText = totalOrders > 0 ? ((premiumOrdersCount / totalOrders) * 100).toFixed(1) + '%' : '0%';

            let bestDay = '-', bestDayRev = 0;
            Object.keys(dailyMap).forEach(d => {
                if (dailyMap[d] > bestDayRev) { bestDay = d; bestDayRev = dailyMap[d]; }
            });
            document.getElementById('wal-best-day').innerText = bestDay;
            document.getElementById('wal-best-day-rev').innerText = formatPLN(bestDayRev);

            const sortedDays = Object.keys(dailyMap).sort();
            walChartDaily.data.labels = sortedDays;
            walChartDaily.data.datasets[0].data = sortedDays.map(k => dailyMap[k]);
            walChartDaily.update();

            let cumSum = 0;
            const cumData = sortedDays.map(d => { cumSum += dailyMap[d]; return cumSum; });
            walChartCumulative.data.labels = sortedDays;
            walChartCumulative.data.datasets[0].data = cumData;
            walChartCumulative.update();

            const sortedMonths = Object.keys(monthlyMap).sort();
            walChartMonthly.data.labels = sortedMonths.map(formatMonth);
            walChartMonthly.data.datasets[0].data = sortedMonths.map(k => monthlyMap[k]);
            walChartMonthly.update();

            walChartPriceDist.data.datasets[0].data = [p299, p399, p254, p339, pOther];
            walChartPriceDist.update();

            walChartDayOfWeek.data.datasets[0].data = dayOfWeekMap;
            walChartDayOfWeek.update();
        }


        // ----------------------------------------------------
        // DATE RANGES & EVENT LISTENERS
        // ----------------------------------------------------
        const allDates = rawTransactions.map(t => t.d).sort();
        const minDateStr = allDates[0];
        const maxDateStr = allDates[allDates.length - 1];

        ['lab', 'szk', 'mos', 'bar', 'wal'].forEach(prefix => {
            const sInput = document.getElementById(prefix + 'StartDate');
            const eInput = document.getElementById(prefix + 'EndDate');
            if (sInput && eInput) {
                sInput.value = minDateStr;
                eInput.value = maxDateStr;
            }
        });

        function setupQuickButtons(prefix, updateFn) {
            const sInput = document.getElementById(prefix + 'StartDate');
            const eInput = document.getElementById(prefix + 'EndDate');

            ['14d', '30d', '90d'].forEach(days => {
                const btn = document.getElementById(`${prefix}-btn-${days}`);
                if (btn) {
                    btn.addEventListener('click', (e) => {
                        const paneId = prefix === 'lab' ? 'labirynt' : prefix === 'szk' ? 'szkolenia' : prefix === 'mos' ? 'mosak' : prefix === 'bar' ? 'debarbaro' : 'walkiewicz';
                        document.querySelectorAll(`#pane-${paneId} .filter-btn`).forEach(b => b.classList.remove('active'));
                        e.target.classList.add('active');

                        let end = new Date(maxDateStr + 'T00:00:00');
                        let start = new Date(maxDateStr + 'T00:00:00');
                        start.setDate(end.getDate() - parseInt(days));
                        sInput.value = start.toISOString().split('T')[0];
                        eInput.value = maxDateStr;
                        updateFn();
                    });
                }
            });

            const btnAll = document.getElementById(`${prefix}-btn-all`);
            if (btnAll) {
                btnAll.addEventListener('click', (e) => {
                    const paneId = prefix === 'lab' ? 'labirynt' : prefix === 'szk' ? 'szkolenia' : prefix === 'mos' ? 'mosak' : prefix === 'bar' ? 'debarbaro' : 'walkiewicz';
                    document.querySelectorAll(`#pane-${paneId} .filter-btn`).forEach(b => b.classList.remove('active'));
                    e.target.classList.add('active');
                    sInput.value = minDateStr;
                    eInput.value = maxDateStr;
                    updateFn();
                });
            }

            sInput.addEventListener('change', () => updateFn());
            eInput.addEventListener('change', () => updateFn());
        }

        setupQuickButtons('lab', updateLabiryntDashboard);
        setupQuickButtons('szk', updateSzkoleniaDashboard);
        setupQuickButtons('mos', updateMosakDashboard);
        setupQuickButtons('bar', updateDebarbaroDashboard);
        setupQuickButtons('wal', updateWalkiewiczDashboard);

        document.getElementById('labVariantSelect').addEventListener('change', updateLabiryntDashboard);
        document.getElementById('szkAuthorSelect').addEventListener('change', updateSzkoleniaDashboard);

        // TAB SWITCHING LOGIC
        function switchTab(tabId) {
            document.querySelectorAll('.nav-tab-btn').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.tab === tabId);
            });
            document.querySelectorAll('.tab-pane').forEach(pane => {
                pane.classList.toggle('active', pane.id === 'pane-' + tabId);
            });

            if (tabId === 'labirynt') {
                updateLabiryntDashboard();
                setTimeout(() => { [labChartDaily, labChartCumulative, labChartMonthly, labChartWeekly, labChartDayOfWeek].forEach(c => c && c.resize()); }, 50);
            } else if (tabId === 'szkolenia') {
                updateSzkoleniaDashboard();
                setTimeout(() => { [szkChartDaily, szkChartCumulative, szkChartAuthors, szkChartTopCourses, szkChartDayOfWeek].forEach(c => c && c.resize()); }, 50);
            } else if (tabId === 'mosak') {
                updateMosakDashboard();
                setTimeout(() => { [mosChartDaily, mosChartDayOfWeek].forEach(c => c && c.resize()); }, 50);
            } else if (tabId === 'debarbaro') {
                updateDebarbaroDashboard();
                setTimeout(() => { [barChartDaily, barChartCumulative, barChartMonthly, barChartPriceDist, barChartDayOfWeek].forEach(c => c && c.resize()); }, 50);
            } else if (tabId === 'walkiewicz') {
                updateWalkiewiczDashboard();
                setTimeout(() => { [walChartDaily, walChartCumulative, walChartMonthly, walChartPriceDist, walChartDayOfWeek].forEach(c => c && c.resize()); }, 50);
            }
        }

        document.querySelectorAll('.nav-tab-btn').forEach(btn => {
            btn.addEventListener('click', () => switchTab(btn.dataset.tab));
        });

        // Initialize default tabs
        updateLabiryntDashboard();
        updateSzkoleniaDashboard();
        updateMosakDashboard();
        updateDebarbaroDashboard();
        updateWalkiewiczDashboard();
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
    git_status = subprocess.run(['git', 'rev-parse', '--is-inside-work-tree'], capture_output=True, text=True)
    if git_status.returncode == 0:
        print("\nWykryto repozytorium Git. Publikuję raport na GitHub Pages...")
        
        subprocess.run(['git', 'add', 'index.html', 'Raport_Labirynt_Euphire.html', '.gitignore', 'generate_report.py', 'Specyfikacja_Analityczna_Labirynt.md', 'README.md', 'Specyfikacja_Wdrozeniowa_Agenta.md'], check=True)
        
        commit_check = subprocess.run(['git', 'status', '--porcelain'], capture_output=True, text=True)
        if commit_check.stdout.strip():
            commit_msg = f"Add dedicated tabs for Prof. de Barbaro (Dobry Terapeuta) and Jacek Walkiewicz (Pelna MOC) - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            subprocess.run(['git', 'commit', '-m', commit_msg], check=True)
            print("Zatwierdzono nowe zmiany w Git.")
            
            remote_check = subprocess.run(['git', 'remote', 'get-url', 'origin'], capture_output=True, text=True)
            if remote_check.returncode == 0:
                push_res = subprocess.run(['git', 'push', 'origin', 'main'], capture_output=True, text=True)
                if push_res.returncode == 0:
                    print("[SUKCES] Zmiany zostały wysłane do repozytorium GitHub!")
                    print(f"Link do raportu dla marketera: https://baciok91.github.io/euphire-raporty-sprzedazy/")
                else:
                    print(f"[BŁĄD] Nie udało się wysłać zmian do GitHub origin: {push_res.stderr.strip()}")
            else:
                print("[INFO] Brak zdalnego repozytorium origin.")
        else:
            print("Brak nowych zmian w raporcie do wysłania.")
except Exception as e:
    print(f"Błąd podczas automatyzacji Git: {e}")
