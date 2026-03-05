# -*- coding: utf-8 -*-
"""
Blocco 3: Pipeline Completa
Python per Excel — Turin Wealth Advisory

Script completo, protezione, report multi-foglio, export PDF,
confronto simulazione vs realta.

Esegui dalla cartella lezione/:
    python 03_pipeline_completa.py
"""

# ============================================================
# SETUP
# ============================================================

import xlwings as xw
import pandas as pd
import numpy as np
import yfinance as yf
import sys
import os
import time
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))

from tw_config import COLORS, AZIENDE, BENCHMARK, NUMBER_FORMATS, CLIENTE
from tw_utils import (
    set_formula, fmt_header, fmt_title, protect_sheet,
    create_workbook, save_and_close, write_table, add_sheet,
    hide_sheet, autofit_all
)

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dati_cache')
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
os.makedirs(CACHE_DIR, exist_ok=True)


def scarica_o_cache(ticker, period='5y'):
    """Scarica dati da Yahoo Finance o usa la cache locale."""
    try:
        df = yf.download(ticker, period=period, progress=False)
        if len(df) > 0:
            nome_file = ticker.replace('^', '').replace('.', '_') + '.csv'
            df.to_csv(os.path.join(CACHE_DIR, nome_file))
            return df
    except Exception:
        pass
    nome_file = ticker.replace('^', '').replace('.', '_') + '.csv'
    path_cache = os.path.join(CACHE_DIR, nome_file)
    if os.path.exists(path_cache):
        print(f'  [cache] {ticker}')
        return pd.read_csv(path_cache, index_col=0, parse_dates=True)
    raise FileNotFoundError(f'Nessun dato per {ticker}')


print(f'Setup completato. Aziende: {[az["nome"] for az in AZIENDE]}')


# ============================================================
# ES10: Funzioni modulari
# ============================================================

print('\n' + '='*60)
print('ES10: Lo script completo')
print('='*60)


def scarica_dati_aziende():
    """Scarica prezzi Close per le 5 aziende Turin Wealth."""
    dati = {}
    for az in AZIENDE:
        df = scarica_o_cache(az['ticker'])
        if isinstance(df.columns, pd.MultiIndex):
            dati[az['nome']] = df['Close'][az['ticker']]
        else:
            dati[az['nome']] = df['Close']
        print(f"  {az['nome']}: {len(df)} giorni")
    return pd.DataFrame(dati)


def scarica_benchmark():
    """Scarica dati degli indici benchmark."""
    dati = {}
    for b in BENCHMARK:
        try:
            df = scarica_o_cache(b['ticker'])
            if isinstance(df.columns, pd.MultiIndex):
                dati[b['nome']] = df['Close'][b['ticker']]
            else:
                dati[b['nome']] = df['Close']
            print(f"  {b['nome']}: {len(df)} giorni")
        except FileNotFoundError as e:
            print(f"  [SKIP] {b['nome']}: {e}")
    return pd.DataFrame(dati)


def calcola_statistiche(prezzi):
    """Calcola statistiche di performance per ogni titolo."""
    rendimenti = prezzi.pct_change().dropna()
    stats = pd.DataFrame({
        'Ultimo prezzo':       prezzi.iloc[-1],
        'Rend. 1M (%)':        ((prezzi.iloc[-1] / prezzi.iloc[-22]) - 1) * 100,
        'Rend. 1Y (%)':        ((prezzi.iloc[-1] / prezzi.iloc[-252]) - 1) * 100,
        'Volatilita ann. (%)': rendimenti.std() * np.sqrt(252) * 100,
        'Sharpe':              (rendimenti.mean() * 252) / (rendimenti.std() * np.sqrt(252)),
    }).round(2)
    return stats


print('Scaricamento dati aziende...')
prezzi = scarica_dati_aziende()
print(f'\nPrezzi: {prezzi.shape[0]} righe x {prezzi.shape[1]} colonne')
print(f'Periodo: {prezzi.index[0].date()} -> {prezzi.index[-1].date()}')

stats = calcola_statistiche(prezzi)
print('\nStatistiche:')
print(stats.to_string())

print('\nScaricamento benchmark...')
bench = scarica_benchmark()
print(f'Benchmark: {list(bench.columns)}')


# ============================================================
# ES11: Protezione e consegna
# ============================================================

print('\n' + '='*60)
print('ES11: Protezione e consegna')
print('='*60)

app = xw.App(visible=True)
time.sleep(0.5)
wb_demo = app.books.add()
ws = wb_demo.sheets[0]
ws.name = 'Report'

# Tabella di esempio
ws.range('A1').value = 'REPORT TRIMESTRALE - FAMIGLIA BIANCHI'
fmt_title(ws.range('A1:D1'))
ws.range('A1:D1').merge()

ws.range('A3').value = [['Titolo', 'Prezzo', 'Variazione %']]
fmt_header(ws.range('A3:C3'))

ws.range('A4').value = [
    ['Terna',     7.85,   0.032],
    ['Ferrari',   420.50, 0.125],
    ['Microsoft', 425.00, 0.087],
]

# Formula nascosta
ws.range('D3').value = 'Var. media'
set_formula(ws.range('D4'), '=MEDIA(C4:C6)')
ws.range('D4').api.FormulaHidden = True

# Sblocca celle modificabili
ws.range('C4:C6').api.Locked = False
protect_sheet(ws, 'password123')

print('Foglio Report protetto')
print(f'  Locked A1: {ws.range("A1").api.Locked}')
print(f'  Locked C4: {ws.range("C4").api.Locked}')
print(f'  FormulaHidden D4: {ws.range("D4").api.FormulaHidden}')

# Foglio very_hidden
ws_segreto = wb_demo.sheets.add('_CORREZIONE')
ws_segreto.range('A1').value = 'Risposte corrette - solo docente!'
ws_segreto.range('A2').value = 'Terna rendimento: +3.2%'
ws_segreto.range('A3').value = 'Ferrari rendimento: +12.5%'
hide_sheet(ws_segreto, very_hidden=True)

fogli_visibili = [s.name for s in wb_demo.sheets if s.api.Visible == -1]
fogli_nascosti = [s.name for s in wb_demo.sheets if s.api.Visible != -1]
print(f'Fogli visibili: {fogli_visibili}')
print(f'Fogli nascosti (very hidden): {fogli_nascosti}')

# Foglio Budget con protezione
ws_budget = wb_demo.sheets.add('Budget')
ws_budget.range('A1').value = 'BUDGET TRIMESTRALE'
fmt_title(ws_budget.range('A1:B1'))
ws_budget.range('A1:B1').merge()

ws_budget.range('A3').value = [['Voce', 'Importo']]
fmt_header(ws_budget.range('A3:B3'))

ws_budget.range('A4').value = [
    ['Commissioni advisory', 14000],
    ['Spese di custodia',     2800],
    ['Altri costi',           1200],
]

ws_budget.range('A7').value = 'TOTALE'
ws_budget.range('A7').font.bold = True
set_formula(ws_budget.range('B7'), '=SOMMA(B4:B6)')
ws_budget.range('B7').api.FormulaHidden = True
ws_budget.range('B7').api.NumberFormatLocal = '#.##0 \u20ac'

ws_budget.range('B4:B6').api.Locked = False
ws_budget.range('B4:B6').api.NumberFormatLocal = '#.##0 \u20ac'
protect_sheet(ws_budget, 'turin2024')

print(f'Foglio Budget protetto:')
print(f'  Locked A1: {ws_budget.range("A1").api.Locked}')
print(f'  Locked B4: {ws_budget.range("B4").api.Locked}')
print(f'  FormulaHidden B7: {ws_budget.range("B7").api.FormulaHidden}')

wb_demo.close()
app.quit()
print('Workbook demo chiuso.')


# ============================================================
# ES12: Report trimestrale Bianchi
# ============================================================

print('\n' + '='*60)
print('ES12: Report trimestrale Bianchi')
print('='*60)


def crea_report_bianchi(prezzi, stats):
    """Genera il report trimestrale completo per la Famiglia Bianchi."""
    wb = create_workbook(visible=True)
    time.sleep(0.5)

    # --- Foglio 1: Copertina ---
    ws_cop = wb.sheets[0]
    ws_cop.name = 'Copertina'
    ws_cop.range('A:A').column_width = 4
    ws_cop.range('B:B').column_width = 28
    ws_cop.range('C:C').column_width = 20

    ws_cop.range('B3').value = 'TURIN WEALTH ADVISORY'
    ws_cop.range('B3').font.size = 22
    ws_cop.range('B3').font.bold = True
    ws_cop.range('B3').font.color = COLORS['header']

    ws_cop.range('B4').value = 'Wealth Management & Advisory'
    ws_cop.range('B4').font.size = 13
    ws_cop.range('B4').font.italic = True
    ws_cop.range('B4').font.color = COLORS['subheader']

    ws_cop.range('B6:D6').api.Borders(9).Weight = 2
    ws_cop.range('B6:D6').api.Borders(9).Color = sum(
        v * (256 ** i) for i, v in enumerate(reversed(COLORS['accent']))
    )

    ws_cop.range('B8').value = 'Report Trimestrale'
    ws_cop.range('B8').font.size = 17
    ws_cop.range('B8').font.bold = True

    trimestre = f"Q{((datetime.now().month - 1) // 3) + 1} {datetime.now().year}"
    ws_cop.range('B9').value = trimestre
    ws_cop.range('B9').font.size = 14
    ws_cop.range('B9').font.color = COLORS['accent']

    ws_cop.range('B12').value = 'Cliente:'
    ws_cop.range('B12').font.bold = True
    ws_cop.range('C12').value = CLIENTE['nome']
    ws_cop.range('B13').value = 'Data:'
    ws_cop.range('B13').font.bold = True
    ws_cop.range('C13').value = datetime.now().strftime('%d/%m/%Y')
    ws_cop.range('B14').value = 'Consulente:'
    ws_cop.range('B14').font.bold = True
    ws_cop.range('C14').value = CLIENTE['advisor']
    ws_cop.range('B15').value = 'Patrimonio gestito:'
    ws_cop.range('B15').font.bold = True
    ws_cop.range('C15').value = CLIENTE['patrimonio_totale']
    ws_cop.range('C15').api.NumberFormatLocal = '\u20ac #.##0'

    # --- Foglio 2: Statistiche ---
    ws_stats = add_sheet(wb, 'Statistiche')
    ws_stats.range('A1').value = 'ANALISI TITOLI \u2014 ' + trimestre
    fmt_title(ws_stats.range('A1:F1'))
    ws_stats.range('A1:F1').merge()

    headers = list(stats.columns)
    data = [[idx] + list(row) for idx, row in stats.iterrows()]
    write_table(ws_stats, (2, 1), ['Titolo'] + headers, data)

    n_rows = len(data)
    ws_stats.range(f'B3:B{2 + n_rows}').api.NumberFormatLocal = '#.##0,00 \u20ac'
    ws_stats.range(f'C3:E{2 + n_rows}').api.NumberFormatLocal = '0,00'
    ws_stats.range(f'F3:F{2 + n_rows}').api.NumberFormatLocal = '0,00'

    # --- Foglio 3: Grafici ---
    ws_graf = add_sheet(wb, 'Grafici')
    ws_graf.range('A1').value = 'ANDAMENTO PREZZI \u2014 Ultimi 60 giorni'
    fmt_title(ws_graf.range('A1:F1'))
    ws_graf.range('A1:F1').merge()

    ultimi_60 = prezzi.tail(60).copy()
    ultimi_60_norm = (ultimi_60 / ultimi_60.iloc[0]) * 100
    ws_graf.range('A3').options(index=True, header=True).value = ultimi_60_norm
    ws_graf.range('A3').api.NumberFormatLocal = 'GG/MM/AAAA'

    last_row = 3 + len(ultimi_60_norm)
    top_pos = ws_graf.range(f'A{last_row + 2}').top
    chart = ws_graf.charts.add(left=10, top=top_pos, width=700, height=380)
    chart.chart_type = 'line'
    chart.set_source_data(ws_graf.range(f'A3:F{last_row}'))
    chart.api[1].HasTitle = True
    chart.api[1].ChartTitle.Text = 'Performance Comparata (base 100)'
    chart.api[1].HasLegend = True
    chart.api[1].Legend.Position = -4107

    autofit_all(wb)
    return wb


print('Generazione report trimestrale...')
report = crea_report_bianchi(prezzi, stats)
print(f'Report: {[s.name for s in report.sheets]}')

# Foglio Correlazione
rendimenti = prezzi.pct_change().dropna()
corr = rendimenti.corr().round(2)

ws_corr = add_sheet(report, 'Correlazione')
ws_corr.range('A1').value = 'MATRICE DI CORRELAZIONE \u2014 Rendimenti Giornalieri'
fmt_title(ws_corr.range('A1:F1'))
ws_corr.range('A1:F1').merge()

ws_corr.range('A3').options(index=True, header=True).value = corr
n_titoli = len(corr.columns)
fmt_header(ws_corr.range(f'A3:{chr(65 + n_titoli)}3'))


def corr_color(val):
    if val >= 1.0:
        return COLORS['locked_bg']
    elif val > 0.8:
        return COLORS['correct_text']
    elif val > 0.5:
        return COLORS['correct_bg']
    elif val > 0.3:
        return COLORS['warning_bg']
    else:
        return COLORS['wrong_bg']


for i, titolo1 in enumerate(corr.index):
    for j, titolo2 in enumerate(corr.columns):
        val = corr.loc[titolo1, titolo2]
        cella = ws_corr.range((4 + i, 2 + j))
        cella.api.NumberFormatLocal = '0,00'
        cella.color = corr_color(val)
        if val > 0.8 and val < 1.0:
            cella.font.color = COLORS['header_text']

autofit_all(report)
print(f'Foglio Correlazione aggiunto. Fogli: {[s.name for s in report.sheets]}')


# ============================================================
# ES13: Esporta PDF
# ============================================================

print('\n' + '='*60)
print('ES13: Esporta PDF')
print('='*60)

# Layout di stampa
for ws in report.sheets:
    ps = ws.api.PageSetup
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.CenterHorizontally = True
    ps.LeftMargin = 36
    ps.RightMargin = 36
    ps.TopMargin = 50
    ps.BottomMargin = 50
    ps.LeftHeader = '&"Calibri,Bold"Turin Wealth Advisory'
    ps.RightHeader = '&D'
    ps.CenterFooter = 'Pagina &P di &N'
    ps.LeftFooter = 'Riservato \u2014 Famiglia Bianchi'

print('Layout di stampa configurato.')

# PDF completo
pdf_path = os.path.join(OUTPUT_DIR, 'Report_Bianchi_Completo.pdf')
report.api.ExportAsFixedFormat(
    Type=0, Filename=pdf_path, Quality=0,
    IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False
)
print(f'PDF completo: {pdf_path}')
if os.path.exists(pdf_path):
    print(f'  Dimensione: {os.path.getsize(pdf_path)/1024:.1f} KB')

# PDF parziale (solo Statistiche e Grafici)
fogli_da_escludere = ['Copertina', 'Correlazione']
pdf_parziale = os.path.join(OUTPUT_DIR, 'Report_Solo_Analisi.pdf')

nascosti = []
try:
    for ws in report.sheets:
        if ws.name in fogli_da_escludere:
            ws.api.Visible = 0
            nascosti.append(ws.name)

    report.api.ExportAsFixedFormat(
        Type=0, Filename=pdf_parziale, Quality=0,
        IgnorePrintAreas=False, OpenAfterPublish=False
    )
    print(f'PDF parziale: {pdf_parziale}')
finally:
    for ws in report.sheets:
        if ws.name in nascosti:
            ws.api.Visible = -1

if os.path.exists(pdf_parziale):
    print(f'  Dimensione: {os.path.getsize(pdf_parziale)/1024:.1f} KB')


# ============================================================
# ES14: Confronto simulazione vs realta
# ============================================================

print('\n' + '='*60)
print('ES14: Il confronto finale')
print('='*60)

path_dati = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'output', 'DATI_Turin_Wealth.xlsx'
)

try:
    sim = pd.read_excel(path_dati, sheet_name='QUOTAZIONI_AZIONI', index_col=0)
    print(f'Dati simulati caricati: {sim.shape}')
    print(f'  Periodo: {sim.index[0]} -> {sim.index[-1]}')
    print(f'  Frequenza: mensile ({len(sim)} osservazioni)')

    print(f'\nDati reali (yfinance):')
    print(f'  Periodo: {prezzi.index[0].date()} -> {prezzi.index[-1].date()}')
    print(f'  Frequenza: giornaliera ({len(prezzi)} osservazioni)')

    # Confronto volatilita
    rend_sim  = sim.pct_change().dropna()
    rend_real = prezzi.pct_change().dropna()
    vol_sim   = rend_sim.std() * np.sqrt(12) * 100
    vol_real  = rend_real.std() * np.sqrt(252) * 100

    print(f'\n  {"Titolo":<15} {"Simul. (%)":>12} {"Reale (%)":>12}')
    print('  ' + '-' * 40)
    for az in AZIENDE:
        nome = az['nome']
        v_s = vol_sim.get(nome, float('nan'))
        v_r = vol_real.get(nome, float('nan'))
        print(f'  {nome:<15} {v_s:>11.1f}% {v_r:>11.1f}%')

    print('\nNote:')
    print('  - I dati simulati usano random.uniform() con seed fisso')
    print('  - I dati reali mostrano cluster di volatilita (ARCH effects)')
    print('  - La simulazione semplice sottostima i rischi di coda')

except FileNotFoundError:
    print('File DATI_Turin_Wealth.xlsx non trovato.')
    print('  Per generarlo: cd scripts && python create_dati.py')
except Exception as e:
    print(f'[ERRORE] {e}')


# ============================================================
# Salva e chiudi
# ============================================================

print('\n' + '='*60)
print('Salvataggio finale')
print('='*60)

output_path = os.path.join(OUTPUT_DIR, 'Report_Trimestrale_Bianchi.xlsx')

try:
    report.save(output_path)
    print(f'Report salvato: {output_path}')
    if os.path.exists(output_path):
        print(f'  Dimensione: {os.path.getsize(output_path)/1024:.1f} KB')
except Exception as e:
    print(f'[ERRORE] Salvataggio fallito: {e}')
    raise

try:
    report.close()
    report.app.quit()
    print('Excel chiuso.')
except Exception:
    print('[INFO] Excel gia chiuso.')

print(f'\nFile generati in {OUTPUT_DIR}:')
for f in os.listdir(OUTPUT_DIR):
    if f.endswith(('.xlsx', '.pdf')):
        print(f'  {f}  ({os.path.getsize(os.path.join(OUTPUT_DIR, f))/1024:.1f} KB)')

print('\nBlocco 3 completato!')
