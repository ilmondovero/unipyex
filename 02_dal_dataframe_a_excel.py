# -*- coding: utf-8 -*-
"""
Blocco 2: Dal DataFrame a Excel
Python per Excel — Turin Wealth Advisory

Esporta dati in Excel con xlwings: formattazione, formule italiane, grafici.

Esegui dalla cartella lezione/:
    python 02_dal_dataframe_a_excel.py
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

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))
from tw_config import COLORS, AZIENDE, NUMBER_FORMATS
from tw_utils import set_formula, fmt_header, protect_sheet

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dati_cache')
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


def scarica_o_cache(ticker, period='5y'):
    """Scarica dati da Yahoo Finance, con fallback su cache locale."""
    try:
        df = yf.download(ticker, period=period, progress=False)
        if len(df) > 0:
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            return df
    except Exception:
        pass
    nome_file = ticker.replace('^', '').replace('.', '_') + '.csv'
    path_cache = os.path.join(CACHE_DIR, nome_file)
    if os.path.exists(path_cache):
        print(f'  Uso cache locale per {ticker}')
        return pd.read_csv(path_cache, index_col=0, parse_dates=True)
    raise FileNotFoundError(f'Nessun dato per {ticker}')


# Carica dati azioni
tickers = [az['ticker'] for az in AZIENDE]
nomi = [az['nome'] for az in AZIENDE]

dati_list = []
for t in tickers:
    df = scarica_o_cache(t)
    dati_list.append(df['Close'].rename(t))

prezzi = pd.concat(dati_list, axis=1)
prezzi.columns = nomi

print(f'Dati caricati: {len(prezzi)} giorni x {len(prezzi.columns)} titoli')
print(f'Periodo: {prezzi.index[0].date()} -> {prezzi.index[-1].date()}')


# ============================================================
# ES06: Da pandas a Excel
# ============================================================

print('\n' + '='*60)
print('ES06: Da pandas a Excel')
print('='*60)

app = xw.App(visible=True)
time.sleep(0.5)
wb = app.books.add()
ws = wb.sheets[0]
ws.name = 'Prezzi'

# Titolo
ws.range('A1').value = 'QUOTAZIONI STORICHE - Turin Wealth Advisory'
ws.range('A1').font.size = 14
ws.range('A1').font.bold = True

# Scrivi ultimi 60 giorni di prezzi
ultimi_60 = prezzi.tail(60)
ws.range('A3').options(index=True, header=True).value = ultimi_60
print(f'Scritte {len(ultimi_60)} righe x {len(ultimi_60.columns)} colonne')

# Leggi i dati indietro da Excel a pandas
dati_letti = ws.range('A3').options(pd.DataFrame, expand='table').value
print(f'Riletti da Excel: {dati_letti.shape}')

# Foglio Rendimenti
rendimenti = prezzi.pct_change().dropna()
ws_rend = wb.sheets.add('Rendimenti')
ws_rend.range('A1').value = 'RENDIMENTI GIORNALIERI - Turin Wealth Advisory'
ws_rend.range('A1').font.size = 14
ws_rend.range('A1').font.bold = True
ws_rend.range('A3').options(index=True, header=True).value = rendimenti
print(f'Foglio Rendimenti: {len(rendimenti)} righe')


# ============================================================
# ES07: Formattazione professionale
# ============================================================

print('\n' + '='*60)
print('ES07: Formattazione professionale')
print('='*60)

ws = wb.sheets['Prezzi']

HEADER_COLOR = COLORS['header']
ACCENT = COLORS['accent']
GOLD = COLORS['gold']
WHITE = COLORS['header_text']
ALT_ROW = COLORS['table_alt_row']

# Header riga 3
header_range = ws.range('A3').expand('right')
header_range.color = HEADER_COLOR
header_range.font.color = WHITE
header_range.font.bold = True
header_range.font.size = 11

# Formato numeri italiano
ultima_riga = ws.range('A3').expand('down').last_cell.row
data_range = ws.range(f'B4:F{ultima_riga}')
data_range.api.NumberFormatLocal = '#.##0,00 \u20ac'

# Righe alternate
for i in range(4, ultima_riga + 1):
    if i % 2 == 0:
        ws.range(f'A{i}:F{i}').color = ALT_ROW

# Bordi
table = ws.range(f'A3:F{ultima_riga}')
border_color = COLORS['border_gray']
border_color_int = border_color[0] + border_color[1] * 256 + border_color[2] * 65536
for border_id in [7, 8, 9, 10]:
    border = table.api.Borders(border_id)
    border.LineStyle = 1
    border.Weight = 2
    border.Color = border_color_int

ws.autofit()
print(f'Formattazione Prezzi: {ultima_riga - 3} righe')

# Formatta foglio Rendimenti
ws_rend = wb.sheets['Rendimenti']
ultima_riga_rend = ws_rend.range('A3').expand('down').last_cell.row

header_rend = ws_rend.range('A3').expand('right')
header_rend.color = COLORS['header']
header_rend.font.color = COLORS['header_text']
header_rend.font.bold = True
header_rend.font.size = 11

dati_rend = ws_rend.range(f'B4:F{ultima_riga_rend}')
dati_rend.api.NumberFormatLocal = '0,00%'

verde = COLORS['correct_bg']
rosso = COLORS['wrong_bg']
valori = dati_rend.value
for r_idx, riga in enumerate(valori):
    for c_idx, val in enumerate(riga):
        if val is not None and isinstance(val, (int, float)):
            cella = ws_rend.range(f'{chr(66 + c_idx)}{4 + r_idx}')
            cella.color = verde if val > 0 else rosso

ws_rend.autofit()
print(f'Formattazione Rendimenti: {ultima_riga_rend - 3} righe, colori verde/rosso')


# ============================================================
# ES08: Formule italiane da Python
# ============================================================

print('\n' + '='*60)
print('ES08: Formule italiane da Python')
print('='*60)

ws = wb.sheets['Prezzi']
ultima_riga = ws.range('A3').expand('down').last_cell.row

# Dimostrazione problema @ vs set_formula()
ws.range('G3').value = 'Media (con @)'
ws.range('G3').font.bold = True
ws.range('G3').color = COLORS['wrong_bg']
ws.range('G4').api.FormulaLocal = '=MEDIA(B4:F4)'
time.sleep(1)

formula_scritta = ws.range('G4').api.FormulaLocal
print(f"FormulaLocal: '{formula_scritta}' — ha @? {'Si' if '@' in formula_scritta else 'No'}")

ws.range('H3').value = 'Media (corretta)'
ws.range('H3').font.bold = True
ws.range('H3').color = COLORS['correct_bg']
set_formula(ws.range('H4'), '=MEDIA(B4:F4)')
time.sleep(1)

formula_corretta = ws.range('H4').api.FormulaLocal
print(f"set_formula(): '{formula_corretta}' — ha @? {'Si' if '@' in formula_corretta else 'No'}")

# Colonne calcolate con set_formula()
ws.range(f'G3:H{ultima_riga}').clear()

ws.range('G3').value = 'Media'
ws.range('H3').value = 'Min'
ws.range('I3').value = 'Max'
fmt_header(ws.range('G3:I3'))

for row in range(4, ultima_riga + 1):
    set_formula(ws.range(f'G{row}'), f'=MEDIA(B{row}:F{row})')
    set_formula(ws.range(f'H{row}'), f'=MIN(B{row}:F{row})')
    set_formula(ws.range(f'I{row}'), f'=MAX(B{row}:F{row})')

ws.range(f'G4:I{ultima_riga}').api.NumberFormatLocal = '#.##0,00 \u20ac'
ws.autofit()
print(f'Colonne Media/Min/Max aggiunte (righe 4-{ultima_riga})')

# Colonne calcolate foglio Rendimenti
ws_rend = wb.sheets['Rendimenti']
ultima_riga_rend = ws_rend.range('A3').expand('down').last_cell.row

ws_rend.range('G3').value = 'Positivi'
ws_rend.range('H3').value = 'Media'
fmt_header(ws_rend.range('G3:H3'))

for row in range(4, ultima_riga_rend + 1):
    set_formula(ws_rend.range(f'G{row}'), f'=CONTA.SE(B{row}:F{row};">0")')
    set_formula(ws_rend.range(f'H{row}'), f'=MEDIA(B{row}:F{row})')

ws_rend.range(f'G4:G{ultima_riga_rend}').api.NumberFormatLocal = '0'
ws_rend.range(f'H4:H{ultima_riga_rend}').api.NumberFormatLocal = '0,00%'
ws_rend.autofit()

formula_g4 = ws_rend.range('G4').api.FormulaLocal
print(f'Formula Positivi (G4): {formula_g4}')
print(f'Presenza @: {"SI (problema!)" if "@" in formula_g4 else "NO (corretto!)"}')


# ============================================================
# ES09: Il grafico che parla
# ============================================================

print('\n' + '='*60)
print('ES09: Il grafico che parla')
print('='*60)

ws = wb.sheets['Prezzi']
ultima_riga = ws.range('A3').expand('down').last_cell.row

# Grafico a linee — andamento prezzi
top_grafico = ws.range(f'A{ultima_riga + 3}').top

chart = ws.charts.add(
    left=ws.range('A1').left,
    top=top_grafico,
    width=650,
    height=370
)
chart.chart_type = 'line'
chart.set_source_data(ws.range(f'A3:F{ultima_riga}'))

chart.api[1].HasTitle = True
chart.api[1].ChartTitle.Text = 'Andamento Prezzi - 5 Titoli Turin Wealth'
chart.api[1].ChartTitle.Font.Size = 13
chart.api[1].ChartTitle.Font.Bold = True

colori_grafico = [
    COLORS['chart_primary'],
    COLORS['chart_secondary'],
    COLORS['chart_accent1'],
    COLORS['chart_accent2'],
    COLORS['chart_accent3'],
]
for i, colore in enumerate(colori_grafico):
    try:
        serie = chart.api[1].SeriesCollection(i + 1)
        colore_int = colore[0] + colore[1] * 256 + colore[2] * 65536
        serie.Format.Line.ForeColor.RGB = colore_int
        serie.Format.Line.Weight = 2
    except Exception:
        pass

print('Grafico linee creato nel foglio Prezzi')

# Grafico a torta — distribuzione
ws_alloc = wb.sheets.add('Allocazione')
nomi_az = list(prezzi.columns)
ultimi_prezzi = prezzi.iloc[-1].values

ws_alloc.range('A1').value = 'Azienda'
ws_alloc.range('B1').value = 'Ultimo Prezzo'
ws_alloc.range('A1:B1').font.bold = True
ws_alloc.range('A1:B1').color = COLORS['header']
ws_alloc.range('A1:B1').font.color = COLORS['header_text']

for i, (nome, prezzo) in enumerate(zip(nomi_az, ultimi_prezzi)):
    ws_alloc.range(f'A{i+2}').value = nome
    ws_alloc.range(f'B{i+2}').value = round(prezzo, 2)

ws_alloc.range('B2:B6').api.NumberFormatLocal = '#.##0,00 \u20ac'
ws_alloc.autofit()

chart_pie = ws_alloc.charts.add(left=220, top=10, width=420, height=320)
chart_pie.chart_type = 'pie'
chart_pie.set_source_data(ws_alloc.range(f'A1:B{len(nomi_az)+1}'))
chart_pie.api[1].HasTitle = True
chart_pie.api[1].ChartTitle.Text = 'Distribuzione per Titolo (ultimo prezzo)'
try:
    chart_pie.api[1].SeriesCollection(1).DataLabels().ShowPercentage = True
    chart_pie.api[1].SeriesCollection(1).DataLabels().ShowValue = False
except Exception:
    pass
print('Grafico a torta creato nel foglio Allocazione')

# Grafico a barre — rendimenti annualizzati
ws_conf = wb.sheets.add('Confronto')
anni = len(prezzi) / 252
rend_annualizzati = []

for nome in nomi_az:
    serie = prezzi[nome].dropna()
    rend_totale = (serie.iloc[-1] / serie.iloc[0]) - 1
    rend_ann = (1 + rend_totale) ** (1 / anni) - 1
    rend_annualizzati.append(rend_ann)

ws_conf.range('A1').value = 'Titolo'
ws_conf.range('B1').value = 'Rendimento Annualizzato'
ws_conf.range('A1:B1').font.bold = True
ws_conf.range('A1:B1').color = COLORS['header']
ws_conf.range('A1:B1').font.color = COLORS['header_text']

for i, (nome, rend) in enumerate(zip(nomi_az, rend_annualizzati)):
    ws_conf.range(f'A{i+2}').value = nome
    ws_conf.range(f'B{i+2}').value = rend

ws_conf.range(f'B2:B{len(nomi_az)+1}').api.NumberFormatLocal = '0,0%'
ws_conf.autofit()

chart_bar = ws_conf.charts.add(left=220, top=10, width=450, height=300)
chart_bar.chart_type = 'bar_clustered'
chart_bar.set_source_data(ws_conf.range(f'A1:B{len(nomi_az)+1}'))
chart_bar.api[1].HasTitle = True
chart_bar.api[1].ChartTitle.Text = 'Rendimento Annualizzato - Turin Wealth Portfolio'

try:
    serie_bar = chart_bar.api[1].SeriesCollection(1)
    for i, rend in enumerate(rend_annualizzati):
        colore = COLORS['chart_accent1'] if rend > 0 else COLORS['chart_accent3']
        colore_int = colore[0] + colore[1] * 256 + colore[2] * 65536
        serie_bar.Points(i + 1).Format.Fill.ForeColor.RGB = colore_int
except Exception:
    pass

print('Grafico rendimenti creato nel foglio Confronto')
for nome, rend in zip(nomi_az, rend_annualizzati):
    print(f'  {nome}: {rend:+.1%}')


# ============================================================
# Salva
# ============================================================

output_path = os.path.join(OUTPUT_DIR, 'report_blocco2.xlsx')
wb.save(output_path)
print(f'\nFile salvato: {output_path}')
print(f'Fogli: {[s.name for s in wb.sheets]}')
print('NON chiudiamo Excel: il file sara usato nel Blocco 3!')
