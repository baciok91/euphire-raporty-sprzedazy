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
    <!-- Google Fonts z EUPHIRE Design System -->
    <link href="https://fonts.googleapis.com/css2?family=Merriweather:ital,wght@0,400;0,700;1,400&family=Montserrat:wght@500;600;700&family=Roboto+Mono:wght@400;500&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --primary-bg: #FAFAFA; /* Frost White */
            --header-bg: #004D54; /* Evergreen */
            --text-main: #1F1F1F; /* Obsidian Black */
            --text-header-light: #FAFAFA;
            --accent: #FCAE2F; /* Ember */
            --accent-subtle: #B2CACC; /* Mist */
            --gradient-dark: #002E32; /* Nocturne */
            
            --font-headings: 'Montserrat', sans-serif;
            --font-paragraphs: 'Merriweather', serif;
            --font-data: 'Roboto Mono', monospace;
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
            line-height: 1.5;
            min-height: 100vh;
        }
        
        /* Typography System */
        h1, h2, h3, h4 {
            font-family: var(--font-headings);
            font-weight: 600;
        }
        
        .overline {
            font-family: var(--font-data);
            line-height: 1.5;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            font-size: 0.85rem;
        }
        
        .data-value {
            font-family: var(--font-data);
            font-weight: 500;
        }
        
        .italic-highlight {
            font-style: italic;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem 3rem 2rem;
        }
        
        /* HEADER SECTION */
        .top-header {
            background-color: var(--header-bg);
            padding: 2rem 0;
            margin-bottom: 2.5rem;
            color: var(--text-header-light);
            border-bottom: 5px solid var(--accent);
            position: relative;
            background-image: url('data:image/svg+xml;utf8,<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg"><filter id="noiseFilter"><feTurbulence type="fractalNoise" baseFrequency="0.65" numOctaves="3" stitchTiles="stitch"/></filter><rect width="100%" height="100%" filter="url(%23noiseFilter)" opacity="0.05"/></svg>');
        }
        
        .header-content {
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .logo-container img {
            height: 48px;
            display: block;
        }
        
        .header-text h1 {
            font-size: 2.2rem;
            margin-top: 0.5rem;
        }
        
        .header-text .overline {
            color: var(--accent);
        }
        
        /* FILTER CONTROLS */
        .filter-container {
            background: #ffffff;
            border: 1px solid var(--accent-subtle);
            border-radius: 8px;
            padding: 1.2rem 1.8rem;
            margin-bottom: 2.5rem;
            display: flex;
            flex-wrap: wrap;
            gap: 1.5rem;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 4px 10px rgba(0, 77, 84, 0.03);
            border-top: 4px solid var(--header-bg);
        }
        
        .filter-group {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            flex-wrap: wrap;
        }
        
        .filter-label {
            font-family: var(--font-headings);
            font-weight: 600;
            color: var(--header-bg);
            font-size: 0.95rem;
        }
        
        .filter-input {
            font-family: var(--font-data);
            border: 1px solid var(--accent-subtle);
            border-radius: 4px;
            padding: 0.4rem 0.8rem;
            color: var(--text-main);
            outline: none;
            font-size: 0.9rem;
            transition: border-color 0.2s, box-shadow 0.2s;
        }
        
        .filter-input:focus {
            border-color: var(--header-bg);
            box-shadow: 0 0 0 3px rgba(0, 77, 84, 0.1);
        }
        
        .btn-group {
            display: flex;
            gap: 0.5rem;
            flex-wrap: wrap;
        }
        
        .filter-btn {
            background: #ffffff;
            border: 1px solid var(--accent-subtle);
            color: var(--header-bg);
            font-family: var(--font-headings);
            font-weight: 500;
            font-size: 0.85rem;
            padding: 0.5rem 1.2rem;
            border-radius: 4px;
            cursor: pointer;
            transition: all 0.2s;
        }
        
        .filter-btn:hover, .filter-btn.active {
            background: var(--header-bg);
            color: var(--text-header-light);
            border-color: var(--header-bg);
        }
        
        /* SUMMARY GRID */
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 1.5rem;
            margin-bottom: 3rem;
        }
        
        .card {
            background-color: #ffffff;
            border-radius: 8px;
            padding: 1.5rem;
            border: 1px solid var(--accent-subtle);
            box-shadow: 0 4px 6px rgba(0, 77, 84, 0.05);
            border-top: 4px solid var(--header-bg);
        }
        
        .card h3 {
            color: var(--header-bg);
            font-size: 1.1rem;
            margin-bottom: 0.5rem;
        }
        
        .card .value {
            font-size: 2.2rem;
            color: var(--text-main);
            margin-bottom: 0.5rem;
        }
        
        .card .value.accent {
            color: var(--accent);
            text-shadow: 0px 1px 2px rgba(0,0,0,0.1);
        }
        
        .card p {
            font-size: 0.95rem;
            color: #555;
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
            background: var(--header-bg);
            color: var(--text-header-light);
            border-radius: 8px;
            padding: 1.5rem;
            border-bottom: 4px solid var(--accent);
            position: relative;
            overflow: hidden;
        }
        
        .variant-card::after {
            content: '';
            position: absolute;
            top: 0; right: 0; bottom: 0; left: 0;
            background-image: url('data:image/svg+xml;utf8,<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg"><filter id="noiseFilter"><feTurbulence type="fractalNoise" baseFrequency="0.65" numOctaves="3" stitchTiles="stitch"/></filter><rect width="100%" height="100%" filter="url(%23noiseFilter)" opacity="0.04"/></svg>');
            pointer-events: none;
        }
        
        .variant-card h4 {
            color: var(--accent);
            font-family: var(--font-data);
            font-size: 1rem;
            margin-bottom: 1rem;
            text-transform: uppercase;
        }
        
        .variant-stats {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            margin-top: 1rem;
        }
        
        .variant-rev { font-size: 1.8rem; font-family: var(--font-data); font-weight: 600; }
        .variant-count { font-size: 1.1rem; opacity: 0.8; font-family: var(--font-data); }

        /* CHARTS AREA */
        .section-title {
            margin-bottom: 1.5rem;
            font-size: 1.8rem;
            color: var(--header-bg);
            display: flex;
            align-items: center;
            gap: 1rem;
        }
        
        .section-title::after {
            content: "";
            flex: 1;
            height: 1px;
            background-color: var(--accent-subtle);
        }
        
        .charts-grid {
            display: grid;
            grid-template-columns: 1fr;
            gap: 2rem;
            margin-bottom: 3rem;
        }
        
        @media (min-width: 1024px) {
            .charts-grid { grid-template-columns: 2fr 1fr; }
        }

        .chart-container {
            background-color: #ffffff;
            border-radius: 8px;
            padding: 1.5rem;
            border: 1px solid var(--accent-subtle);
            height: 400px;
        }
        
        /* INSIGHTS */
        .insights {
            background-color: #ffffff;
            border-radius: 8px;
            padding: 2.5rem;
            border: 1px solid var(--accent-subtle);
            border-left: 6px solid var(--header-bg);
            margin-bottom: 2rem;
            box-shadow: 0 4px 15px rgba(0,0,0,0.03);
        }
        
        .insights h2 {
            color: var(--header-bg);
            margin-bottom: 1.5rem;
            font-size: 1.6rem;
        }
        
        .insights ul {
            list-style: none;
            margin-top: 1rem;
        }
        
        .insights li {
            margin-bottom: 1rem;
            padding-left: 1.5rem;
            position: relative;
            font-size: 1.1rem;
        }
        
        .insights li::before {
            content: "•";
            color: var(--accent);
            font-size: 2rem;
            position: absolute;
            left: 0;
            top: -0.5rem;
        }
        
        .badge {
            background: var(--header-bg);
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-family: var(--font-data);
            font-size: 0.9rem;
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
                <!-- EUPHIRE Primary Logo (Yellow Ember on Evergreen context) -->
                <img src="https://euphire.pl/wp-content/uploads/2025/11/logo_primary.svg" alt="Euphire Logo">
            </div>
        </div>
    </div>

    <div class="container">
        
        <!-- DYNAMIC FILTER CONTAINER -->
        <div class="filter-container">
            <div class="filter-group">
                <span class="filter-label">Filtrowanie zakresu:</span>
                <input type="date" id="startDate" class="filter-input">
                <span class="filter-label" style="font-weight: normal; color: #555;">do</span>
                <input type="date" id="endDate" class="filter-input">
            </div>
            <div class="btn-group">
                <button class="filter-btn" id="btn-14d">Ostatnie 14 dni</button>
                <button class="filter-btn" id="btn-30d">Ostatnie 30 dni</button>
                <button class="filter-btn active" id="btn-all">Cały okres</button>
            </div>
        </div>
        
        <h2 class="section-title">Podsumowanie Okresu</h2>
        <div class="summary-grid">
            <div class="card">
                <h3>Całkowity Przychód</h3>
                <div class="value accent data-value" id="total-revenue">0,00 zł</div>
                <p>Tylko przychód z przypisanych wariantów</p>
            </div>
            <div class="card">
                <h3>Liczba Zamówień</h3>
                <div class="value data-value" id="total-orders">0</div>
                <p>Potwierdzone transakcje w zbiorze</p>
            </div>
            <div class="card">
                <h3>Łącznie Sprzedano (Sztuk)</h3>
                <div class="value data-value" id="total-sales">0</div>
                <p>Obejmuje sztuki wszystkich wariantów</p>
            </div>
            <div class="card">
                <h3>Średni Wartość Zam. (AOV)</h3>
                <div class="value data-value" id="avg-order-value">0,00 zł</div>
                <p>Średni koszyk z zakupem wariantu</p>
            </div>
        </div>
        
        <h2 class="section-title">Wyniki Koszykowe Na Warianty</h2>
        <div class="variant-grid">
            <div class="variant-card" id="card-fizyczna">
                <h4>Wersja Fizyczna (Pudełko)</h4>
                <div class="variant-stats">
                    <div class="variant-rev" id="rev-fizyczna">0,00 zł</div>
                    <div class="variant-count" id="count-fizyczna">0 szt.</div>
                </div>
            </div>
            <div class="variant-card" id="card-cyfrowa">
                <h4>Wersja Cyfrowa (On-line)</h4>
                <div class="variant-stats">
                    <div class="variant-rev" id="rev-cyfrowa">0,00 zł</div>
                    <div class="variant-count" id="count-cyfrowa">0 szt.</div>
                </div>
            </div>
            <div class="variant-card" id="card-pakiet">
                <h4>Pakiet (Pudełko + On-line)</h4>
                <div class="variant-stats">
                    <div class="variant-rev" id="rev-pakiet">0,00 zł</div>
                    <div class="variant-count" id="count-pakiet">0 szt.</div>
                </div>
            </div>
        </div>
        
        <h2 class="section-title">Dynamika Sprzedaży</h2>
        <div class="charts-grid">
            <div class="chart-container">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: var(--header-bg);">Trendy Dzienne</h3>
                <div style="position: relative; height: 320px; width: 100%;">
                    <canvas id="dailyChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3 style="font-family: var(--font-headings); margin-bottom: 1rem; color: var(--header-bg);">Przychód Tygodniowy</h3>
                <div style="position: relative; height: 320px; width: 100%;">
                    <canvas id="weeklyChart"></canvas>
                </div>
            </div>
        </div>
        
        <div class="insights">
            <h2>Wnioski Zgodne z <span class="italic-highlight">Factami, Metodą, Praktyką</span></h2>
            <p>Zautomatyzowana analiza odnotowała następujące momenty skrajne z podzbioru danych dotyczących wyłącznie produktu "Labirynt Rozmów":</p>
            <ul>
                <li><strong>Najsilniejszy dzień pod względem przychodów:</strong> Dzień <span class="badge" id="best-day-badge">-</span> z wynikiem <span class="data-value" id="best-day-rev">0,00 zł</span>.</li>
                <li><strong>Najsilniejszy tydzień łączny:</strong> <span class="badge" id="best-week-badge">-</span> generujący przychód w wysokości <span class="data-value" id="best-week-rev">0,00 zł</span>.</li>
                <li>Zestawienie <span class="italic-highlight">wariantów pudełkowych oraz online</span> wskazuje, jak użytkownicy wybierają formę gry w skali danego ujęcia czasu. Pamiętaj, aby śledzić spadek / przyrost zainteresowania pakietami premium w okresach przedświątecznych lub po kampaniach mailingowych marki Euphire.</li>
            </ul>
        </div>
    </div>

    <script>
        // Data injected via Python
        const rawTransactions = {{RAW_TRANSACTIONS_JSON}};
        
        // Helper to get ISO Week format
        function getISOWeek(dateString) {
            const date = new Date(dateString);
            const tempDate = new Date(date.valueOf());
            tempDate.setDate(tempDate.getDate() + 4 - (tempDate.getDay() || 7));
            const yearStart = new Date(tempDate.getFullYear(), 0, 1);
            const weekNo = Math.ceil((((tempDate - yearStart) / 86400000) + 1) / 7);
            const weekStr = weekNo < 10 ? '0' + weekNo : weekNo;
            return `${tempDate.getFullYear()}-W${weekStr}`;
        }

        Chart.defaults.font.family = "'Roboto Mono', monospace";
        Chart.defaults.color = "#1F1F1F";

        // Daily Chart Render (Line with gradient - Evergreen to Mist)
        const ctxDaily = document.getElementById('dailyChart').getContext('2d');
        const gradientDaily = ctxDaily.createLinearGradient(0, 0, 0, 320);
        gradientDaily.addColorStop(0, 'rgba(0, 77, 84, 0.4)'); // Evergreen 
        gradientDaily.addColorStop(1, 'rgba(178, 202, 204, 0.05)'); // Mist
        
        const dailyChartInstance = new Chart(ctxDaily, {
            type: 'line',
            data: {
                labels: [],
                datasets: [{
                    label: 'Przychód Brutto',
                    data: [],
                    borderColor: '#004D54',
                    backgroundColor: gradientDaily,
                    borderWidth: 3,
                    pointBackgroundColor: '#FAFAFA',
                    pointBorderColor: '#FCAE2F',
                    pointBorderWidth: 3,
                    pointRadius: 4,
                    fill: true,
                    tension: 0.3
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#1F1F1F',
                        titleColor: '#FCAE2F',
                        bodyColor: '#FAFAFA',
                        padding: 12,
                        callbacks: {
                            label: function(ct) { return ct.parsed.y.toLocaleString('pl-PL') + ' zł'; }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(0,0,0,0.05)' },
                        ticks: { callback: function(v) { return v.toLocaleString('pl-PL'); } }
                    },
                    x: { grid: { display: false } }
                }
            }
        });
        
        // Weekly Chart Render (Bar - Ember & Evergreen colors)
        const ctxWeekly = document.getElementById('weeklyChart').getContext('2d');
        const weeklyChartInstance = new Chart(ctxWeekly, {
            type: 'bar',
            data: {
                labels: [],
                datasets: [{
                    label: 'Przychód za Tydzień',
                    data: [],
                    backgroundColor: '#FCAE2F',
                    borderRadius: 4,
                    barPercentage: 0.5
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#004D54',
                        titleColor: '#FCAE2F',
                        bodyColor: '#FAFAFA',
                        padding: 12,
                        callbacks: {
                            label: function(ct) { return ct.parsed.y.toLocaleString('pl-PL') + ' zł'; }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(0,0,0,0.05)' },
                        ticks: { callback: function(v) { return v.toLocaleString('pl-PL'); } }
                    },
                    x: { grid: { display: false } }
                }
            }
        });

        // Dashboard Update Logic
        function updateDashboard(startStr, endStr) {
            const startDate = new Date(startStr + 'T00:00:00');
            const endDate = new Date(endStr + 'T23:59:59');
            
            // Filter
            const filtered = rawTransactions.filter(t => {
                const d = new Date(t.Dzien + 'T00:00:00');
                return d >= startDate && d <= endDate;
            });
            
            // Metrics
            let totalRevenue = 0;
            let totalSales = 0;
            let totalOrders = filtered.length;
            
            const varStats = {
                'Wersja Fizyczna (Pudełko)': { rev: 0, count: 0 },
                'Wersja Cyfrową (On-line)': { rev: 0, count: 0 },
                'Pakiet (Pudełko + On-line)': { rev: 0, count: 0 }
            };
            
            const dailyMap = {};
            const weeklyMap = {};
            
            // Group and aggregate
            filtered.forEach(t => {
                const rev = parseFloat(t['Kwota Brutto']) || 0;
                const qty = parseInt(t['Sztuk']) || 0;
                
                totalRevenue += rev;
                totalSales += qty;
                
                if (varStats[t.Wariant]) {
                    varStats[t.Wariant].rev += rev;
                    varStats[t.Wariant].count += qty;
                }
                
                dailyMap[t.Dzien] = (dailyMap[t.Dzien] || 0) + rev;
                
                const week = getISOWeek(t.Dzien);
                weeklyMap[week] = (weeklyMap[week] || 0) + rev;
            });
            
            const avgOrderValue = totalOrders > 0 ? (totalRevenue / totalOrders) : 0;
            
            // Render General Metrics
            document.getElementById('total-revenue').innerText = totalRevenue.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('total-orders').innerText = totalOrders.toLocaleString('pl-PL');
            document.getElementById('total-sales').innerText = totalSales.toLocaleString('pl-PL');
            document.getElementById('avg-order-value').innerText = avgOrderValue.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            
            // Render Variants
            document.getElementById('rev-fizyczna').innerText = varStats['Wersja Fizyczna (Pudełko)'].rev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('count-fizyczna').innerText = varStats['Wersja Fizyczna (Pudełko)'].count.toLocaleString('pl-PL') + ' szt.';
            
            document.getElementById('rev-cyfrowa').innerText = varStats['Wersja Cyfrową (On-line)'].rev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('count-cyfrowa').innerText = varStats['Wersja Cyfrową (On-line)'].count.toLocaleString('pl-PL') + ' szt.';
            
            document.getElementById('rev-pakiet').innerText = varStats['Pakiet (Pudełko + On-line)'].rev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('count-pakiet').innerText = varStats['Pakiet (Pudełko + On-line)'].count.toLocaleString('pl-PL') + ' szt.';
            
            // Sort daily
            const dailyLabels = Object.keys(dailyMap).sort();
            const dailyRevenue = dailyLabels.map(k => dailyMap[k]);
            
            // Sort weekly
            const weeklyLabels = Object.keys(weeklyMap).sort();
            const weeklyRevenue = weeklyLabels.map(k => weeklyMap[k]);
            
            // Update Daily Chart
            dailyChartInstance.data.labels = dailyLabels;
            dailyChartInstance.data.datasets[0].data = dailyRevenue;
            dailyChartInstance.update();
            
            // Update Weekly Chart
            weeklyChartInstance.data.labels = weeklyLabels;
            weeklyChartInstance.data.datasets[0].data = weeklyRevenue;
            weeklyChartInstance.update();
            
            // Update Insights
            let bestDay = 'Brak';
            let bestDayRev = 0;
            if (dailyRevenue.length > 0) {
                const maxDRev = Math.max(...dailyRevenue);
                bestDayRev = maxDRev;
                const idx = dailyRevenue.indexOf(maxDRev);
                bestDay = dailyLabels[idx];
            }
            
            let bestWeek = 'Brak';
            let bestWeekRev = 0;
            if (weeklyRevenue.length > 0) {
                const maxWRev = Math.max(...weeklyRevenue);
                bestWeekRev = maxWRev;
                const idx = weeklyRevenue.indexOf(maxWRev);
                bestWeek = weeklyLabels[idx];
            }
            
            document.getElementById('best-day-badge').innerText = bestDay;
            document.getElementById('best-day-rev').innerText = bestDayRev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
            document.getElementById('best-week-badge').innerText = bestWeek;
            document.getElementById('best-week-rev').innerText = bestWeekRev.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' zł';
        }

        // Initialize Dates
        let sortedDates = rawTransactions.map(t => t.Dzien).sort();
        let minDateStr = sortedDates.length > 0 ? sortedDates[0] : '';
        let maxDateStr = sortedDates.length > 0 ? sortedDates[sortedDates.length - 1] : '';

        const startInput = document.getElementById('startDate');
        const endInput = document.getElementById('endDate');
        
        startInput.value = minDateStr;
        endInput.value = maxDateStr;
        startInput.min = minDateStr;
        startInput.max = maxDateStr;
        endInput.min = minDateStr;
        endInput.max = maxDateStr;

        // Add event listeners
        startInput.addEventListener('change', () => {
            clearButtonActive();
            updateDashboard(startInput.value, endInput.value);
        });
        
        endInput.addEventListener('change', () => {
            clearButtonActive();
            updateDashboard(startInput.value, endInput.value);
        });
        
        function clearButtonActive() {
            document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.remove('active'));
        }
        
        document.getElementById('btn-14d').addEventListener('click', (e) => {
            clearButtonActive();
            e.target.classList.add('active');
            let end = new Date(maxDateStr);
            let start = new Date(maxDateStr);
            start.setDate(end.getDate() - 14);
            const startStr = start.toISOString().split('T')[0];
            startInput.value = startStr;
            endInput.value = maxDateStr;
            updateDashboard(startStr, maxDateStr);
        });
        
        document.getElementById('btn-30d').addEventListener('click', (e) => {
            clearButtonActive();
            e.target.classList.add('active');
            let end = new Date(maxDateStr);
            let start = new Date(maxDateStr);
            start.setDate(end.getDate() - 30);
            const startStr = start.toISOString().split('T')[0];
            startInput.value = startStr;
            endInput.value = maxDateStr;
            updateDashboard(startStr, maxDateStr);
        });
        
        document.getElementById('btn-all').addEventListener('click', (e) => {
            clearButtonActive();
            e.target.classList.add('active');
            startInput.value = minDateStr;
            endInput.value = maxDateStr;
            updateDashboard(minDateStr, maxDateStr);
        });

        // Initial render
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
        subprocess.run(['git', 'add', 'index.html', 'Raport_Labirynt_Euphire.html', '.gitignore', 'generate_report.py', 'Specyfikacja_Analityczna_Labirynt.md', 'README.md'], check=True)
        
        # 2. Wykonaj commit (jeśli są zmiany)
        commit_check = subprocess.run(['git', 'status', '--porcelain'], capture_output=True, text=True)
        if commit_check.stdout.strip():
            commit_msg = f"Update report with dynamic filters from {os.path.basename(file_path)} - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
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

# --- AUTOMATYCZNE OTWIERANIE W CHROME ---
try:
    print("\nOtwieram wygenerowany raport w Google Chrome...")
    # Otwórz lokalną ścieżkę w Chrome
    local_url = os.path.abspath('index.html')
    subprocess.run(['open', '-a', 'Google Chrome', f'file://{local_url}'])
except Exception as e:
    print(f"Nie udało się otworzyć raportu w Chrome: {e}")
