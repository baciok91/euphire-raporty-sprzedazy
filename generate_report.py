import pandas as pd
import json
import os
from datetime import datetime

file_path = '/Users/maciekckoklormam91/Desktop/Analiza sprzedazy EUPHIE/od wrzesnia do maja.xlsx'

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

# Statystyki ogólne dla Labiryntu
total_revenue = df_filtered['Kwota Brutto'].sum()
total_sales = df_filtered['Sztuk'].sum()
total_orders = len(df_filtered)
avg_order_value = total_revenue / total_orders if total_orders > 0 else 0

# Statystyki per wariant
variant_stats = df_filtered.groupby('Wariant').agg({'Kwota Brutto': 'sum', 'Sztuk': 'sum'}).reset_index()
variant_labels = variant_stats['Wariant'].tolist()
variant_revenue = variant_stats['Kwota Brutto'].tolist()
variant_count = variant_stats['Sztuk'].tolist()

# Przygotowanie danych dziennych dla wykresów
df_filtered['Dzien'] = df_filtered['Data'].dt.date
daily_stats = df_filtered.groupby('Dzien').agg({'Kwota Brutto': 'sum'}).reset_index()
daily_stats['Dzien'] = daily_stats['Dzien'].astype(str)

daily_labels = daily_stats['Dzien'].tolist()
daily_revenue = daily_stats['Kwota Brutto'].tolist()

# Przygotowanie danych tygodniowych (Rok-Tydzień)
df_filtered['Tydzien'] = df_filtered['Data'].dt.strftime('%G-W%V')
weekly_stats = df_filtered.groupby('Tydzien').agg({'Kwota Brutto': 'sum'}).reset_index()
weekly_stats = weekly_stats.sort_values('Tydzien')

weekly_labels = weekly_stats['Tydzien'].tolist()
weekly_revenue = weekly_stats['Kwota Brutto'].tolist()

# Automatyczna analiza
best_day_idx = daily_revenue.index(max(daily_revenue)) if daily_revenue else -1
best_day = daily_labels[best_day_idx] if best_day_idx != -1 else "Brak"
best_day_rev = max(daily_revenue) if daily_revenue else 0

best_week_idx = weekly_revenue.index(max(weekly_revenue)) if weekly_revenue else -1
best_week = weekly_labels[best_week_idx] if best_week_idx != -1 else "Brak"
best_week_rev = max(weekly_revenue) if weekly_revenue else 0

# HTML (EUPHIRE Design System)
html_content = f"""<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>EUPHIRE | Analiza Labirynt Rozmów</title>
    <!-- Google Fonts z EUPHIRE Design System -->
    <link href="https://fonts.googleapis.com/css2?family=Merriweather:ital,wght@0,400;0,700;1,400&family=Montserrat:wght@500;600;700&family=Roboto+Mono:wght@400;500&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {{
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
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            background-color: var(--primary-bg);
            color: var(--text-main);
            font-family: var(--font-paragraphs);
            line-height: 1.5;
            min-height: 100vh;
        }}
        
        /* Typography System */
        h1, h2, h3, h4 {{
            font-family: var(--font-headings);
            font-weight: 600;
        }}
        
        .overline {{
            font-family: var(--font-data);
            line-height: 1.5;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            font-size: 0.85rem;
        }}
        
        .data-value {{
            font-family: var(--font-data);
            font-weight: 500;
        }}
        
        .italic-highlight {{
            font-style: italic;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem 3rem 2rem;
        }}
        
        /* HEADER SECTION */
        .top-header {{
            background-color: var(--header-bg);
            padding: 2rem 0;
            margin-bottom: 3rem;
            color: var(--text-header-light);
            border-bottom: 5px solid var(--accent);
            position: relative;
            background-image: url('data:image/svg+xml;utf8,<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg"><filter id="noiseFilter"><feTurbulence type="fractalNoise" baseFrequency="0.65" numOctaves="3" stitchTiles="stitch"/></filter><rect width="100%" height="100%" filter="url(%23noiseFilter)" opacity="0.05"/></svg>');
        }}
        
        .header-content {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        
        .logo-container img {{
            height: 48px;
            display: block;
        }}
        
        .header-text h1 {{
            font-size: 2.2rem;
            margin-top: 0.5rem;
        }}
        
        .header-text .overline {{
            color: var(--accent);
        }}
        
        /* SUMMARY GRID */
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 1.5rem;
            margin-bottom: 3rem;
        }}
        
        .card {{
            background-color: #ffffff;
            border-radius: 8px;
            padding: 1.5rem;
            border: 1px solid var(--accent-subtle);
            box-shadow: 0 4px 6px rgba(0, 77, 84, 0.05);
            border-top: 4px solid var(--header-bg);
        }}
        
        .card h3 {{
            color: var(--header-bg);
            font-size: 1.1rem;
            margin-bottom: 0.5rem;
        }}
        
        .card .value {{
            font-size: 2.2rem;
            color: var(--text-main);
            margin-bottom: 0.5rem;
        }}
        
        .card .value.accent {{
            color: var(--accent);
            text-shadow: 0px 1px 2px rgba(0,0,0,0.1);
        }}
        
        .card p {{
            font-size: 0.95rem;
            color: #555;
        }}
        
        /* VARIANT ANALYSIS */
        .variant-grid {{
            display: flex;
            gap: 1.5rem;
            margin-bottom: 3rem;
            flex-wrap: wrap;
        }}
        
        .variant-card {{
            flex: 1;
            min-width: 300px;
            background: var(--header-bg);
            color: var(--text-header-light);
            border-radius: 8px;
            padding: 1.5rem;
            border-bottom: 4px solid var(--accent);
            position: relative;
            overflow: hidden;
        }}
        
        .variant-card::after {{
            content: '';
            position: absolute;
            top: 0; right: 0; bottom: 0; left: 0;
            background-image: url('data:image/svg+xml;utf8,<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg"><filter id="noiseFilter"><feTurbulence type="fractalNoise" baseFrequency="0.65" numOctaves="3" stitchTiles="stitch"/></filter><rect width="100%" height="100%" filter="url(%23noiseFilter)" opacity="0.04"/></svg>');
            pointer-events: none;
        }}
        
        .variant-card h4 {{
            color: var(--accent);
            font-family: var(--font-data);
            font-size: 1rem;
            margin-bottom: 1rem;
            text-transform: uppercase;
        }}
        
        .variant-stats {{
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            margin-top: 1rem;
        }}
        
        .variant-rev {{ font-size: 1.8rem; font-family: var(--font-data); font-weight: 600; }}
        .variant-count {{ font-size: 1.1rem; opacity: 0.8; font-family: var(--font-data); }}

        /* CHARTS AREA */
        .section-title {{
            margin-bottom: 1.5rem;
            font-size: 1.8rem;
            color: var(--header-bg);
            display: flex;
            align-items: center;
            gap: 1rem;
        }}
        
        .section-title::after {{
            content: "";
            flex: 1;
            height: 1px;
            background-color: var(--accent-subtle);
        }}
        
        .charts-grid {{
            display: grid;
            grid-template-columns: 1fr;
            gap: 2rem;
            margin-bottom: 3rem;
        }}
        
        @media (min-width: 1024px) {{
            .charts-grid {{ grid-template-columns: 2fr 1fr; }}
        }}

        .chart-container {{
            background-color: #ffffff;
            border-radius: 8px;
            padding: 1.5rem;
            border: 1px solid var(--accent-subtle);
            height: 400px;
        }}
        
        /* INSIGHTS */
        .insights {{
            background-color: #ffffff;
            border-radius: 8px;
            padding: 2.5rem;
            border: 1px solid var(--accent-subtle);
            border-left: 6px solid var(--header-bg);
            margin-bottom: 2rem;
            box-shadow: 0 4px 15px rgba(0,0,0,0.03);
        }}
        
        .insights h2 {{
            color: var(--header-bg);
            margin-bottom: 1.5rem;
            font-size: 1.6rem;
        }}
        
        .insights ul {{
            list-style: none;
            margin-top: 1rem;
        }}
        
        .insights li {{
            margin-bottom: 1rem;
            padding-left: 1.5rem;
            position: relative;
            font-size: 1.1rem;
        }}
        
        .insights li::before {{
            content: "•";
            color: var(--accent);
            font-size: 2rem;
            position: absolute;
            left: 0;
            top: -0.5rem;
        }}
        
        .badge {{
            background: var(--header-bg);
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-family: var(--font-data);
            font-size: 0.9rem;
        }}
    </style>
</head>
<body>
    <div class="top-header">
        <div class="header-content">
            <div class="header-text">
                <span class="overline">Raport Analityczny / Produkt</span>
                <h1>Sprzedaż: Labirynt Rozmów</h1>
                <span class="overline">Wygenerowano: {datetime.now().strftime('%d.%m.%Y %H:%M')}</span>
            </div>
            <div class="logo-container">
                <!-- EUPHIRE Primary Logo (Yellow Ember on Evergreen context) -->
                <img src="https://euphire.pl/wp-content/uploads/2025/11/logo_primary.svg" alt="Euphire Logo">
            </div>
        </div>
    </div>

    <div class="container">
        
        <h2 class="section-title">Podsumowanie Okresu</h2>
        <div class="summary-grid">
            <div class="card">
                <h3>Całkowity Przychód</h3>
                <div class="value accent data-value">{total_revenue:,.2f} zł</div>
                <p>Tylko przychód z przypisanych wariantów</p>
            </div>
            <div class="card">
                <h3>Liczba Zamówień</h3>
                <div class="value data-value">{total_orders}</div>
                <p>Potwierdzone transakcje w zbiorze</p>
            </div>
            <div class="card">
                <h3>Łącznie Sprzedano (Sztuk)</h3>
                <div class="value data-value">{total_sales:,.0f}</div>
                <p>Obejmuje sztuki wszystkich wariantów</p>
            </div>
            <div class="card">
                <h3>Średni Wartość Zam. (AOV)</h3>
                <div class="value data-value">{avg_order_value:,.2f} zł</div>
                <p>Średni koszyk z zakupem wariantu</p>
            </div>
        </div>
        
        <h2 class="section-title">Wyniki Koszykowe Na Warianty</h2>
        <div class="variant-grid">
        """
        
# Add variant specific sections dynamically
for i in range(len(variant_labels)):
    html_content += f"""
            <div class="variant-card">
                <h4>{variant_labels[i]}</h4>
                <div class="variant-stats">
                    <div class="variant-rev">{variant_revenue[i]:,.2f} zł</div>
                    <div class="variant-count">{variant_count[i]:,.0f} szt.</div>
                </div>
            </div>
    """

html_content += f"""
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
                <li><strong>Najsilniejszy dzień pod względem przychodów:</strong> Dzień <span class="badge">{best_day}</span> z wynikiem <span class="data-value">{best_day_rev:,.2f} zł</span>.</li>
                <li><strong>Najsilniejszy tydzień łączny:</strong> <span class="badge">{best_week}</span> generujący przychód w wysokości <span class="data-value">{best_week_rev:,.2f} zł</span>.</li>
                <li>Zestawienie <span class="italic-highlight">wariantów pudełkowych oraz online</span> wskazuje, jak użytkownicy wybierają formę gry w skali danego ujęcia czasu. Pamiętaj, aby śledzić spadek / przyrost zainteresowania pakietami premium w okresach przedświątecznych lub po kampaniach mailingowych marki Euphire.</li>
            </ul>
        </div>
    </div>

    <script>
        // Data injected via Python
        const dailyLabels = {json.dumps(daily_labels)};
        const dailyRevenue = {json.dumps(daily_revenue)};
        const weeklyLabels = {json.dumps(weekly_labels)};
        const weeklyRevenue = {json.dumps(weekly_revenue)};
        
        Chart.defaults.font.family = "'Roboto Mono', monospace";
        Chart.defaults.color = "#1F1F1F";

        // Daily Chart Render (Line with gradient - Evergreen to Mist)
        const ctxDaily = document.getElementById('dailyChart').getContext('2d');
        const gradientDaily = ctxDaily.createLinearGradient(0, 0, 0, 320);
        gradientDaily.addColorStop(0, 'rgba(0, 77, 84, 0.4)'); // Evergreen 
        gradientDaily.addColorStop(1, 'rgba(178, 202, 204, 0.05)'); // Mist
        
        new Chart(ctxDaily, {{
            type: 'line',
            data: {{
                labels: dailyLabels,
                datasets: [{{
                    label: 'Przychód Brutto',
                    data: dailyRevenue,
                    borderColor: '#004D54',
                    backgroundColor: gradientDaily,
                    borderWidth: 3,
                    pointBackgroundColor: '#FAFAFA',
                    pointBorderColor: '#FCAE2F',
                    pointBorderWidth: 3,
                    pointRadius: 5,
                    fill: true,
                    tension: 0.3
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }},
                    tooltip: {{
                        backgroundColor: '#1F1F1F',
                        titleColor: '#FCAE2F',
                        bodyColor: '#FAFAFA',
                        padding: 12,
                        callbacks: {{
                            label: function(ct) {{ return ct.parsed.y.toLocaleString('pl-PL') + ' zł'; }}
                        }}
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        grid: {{ color: 'rgba(0,0,0,0.05)' }},
                        ticks: {{ callback: function(v) {{ return v.toLocaleString('pl-PL'); }} }}
                    }},
                    x: {{ grid: {{ display: false }} }}
                }}
            }}
        }});
        
        // Weekly Chart Render (Bar - Ember & Evergreen colors)
        const ctxWeekly = document.getElementById('weeklyChart').getContext('2d');
        new Chart(ctxWeekly, {{
            type: 'bar',
            data: {{
                labels: weeklyLabels,
                datasets: [{{
                    label: 'Przychód za Tydzień',
                    data: weeklyRevenue,
                    backgroundColor: '#FCAE2F',
                    borderRadius: 4,
                    barPercentage: 0.5
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }},
                    tooltip: {{
                        backgroundColor: '#004D54',
                        titleColor: '#FCAE2F',
                        bodyColor: '#FAFAFA',
                        padding: 12,
                        callbacks: {{
                            label: function(ct) {{ return ct.parsed.y.toLocaleString('pl-PL') + ' zł'; }}
                        }}
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        grid: {{ color: 'rgba(0,0,0,0.05)' }},
                        ticks: {{ callback: function(v) {{ return v.toLocaleString('pl-PL'); }} }}
                    }},
                    x: {{ grid: {{ display: false }} }}
                }}
            }}
        }});
    </script>
</body>
</html>
"""

output_path = '/Users/maciekckoklormam91/Desktop/Analiza sprzedazy EUPHIE/Raport_Labirynt_Euphire.html'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Raport wygenerowany według Euphire Design System: {output_path}")
