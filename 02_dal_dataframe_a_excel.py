# -*- coding: utf-8 -*-
"""
Blocco 2: Dal DataFrame a Excel
Turin Wealth Advisory - Percorso di Esercitazione Excel

Esegui dalla cartella lezione/:
    python 02_dal_dataframe_a_excel.py
"""

# Imposta su True per vedere le soluzioni degli esercizi
MOSTRA_SOLUZIONI = False

# ==============================================================================
# Blocco 2: Dal DataFrame a Excel
#
# Scenario: Il Dott. Ferretti vuole un report Excel professionale per i Bianchi.
#
# In questo blocco impari a:
#   ES06 - Esportare un DataFrame pandas in Excel con xlwings
#   ES07 - Formattare il foglio con i colori e lo stile di Turin Wealth
#   ES08 - Inserire formule italiane da Python (e capire perché e' diverso dall'inglese)
#   ES09 - Creare grafici professionali direttamente da Python
#
# Architettura xlwings:
#   Python --> xlwings --> COM --> Excel (aperto)
#
# xlwings non scrive file .xlsx direttamente: parla con Excel attraverso COM
# (Component Object Model su Windows). Excel deve essere installato.
# ==============================================================================

import xlwings as xw
import pandas as pd
import numpy as np
import yfinance as yf
import sys
import os
import time

# Aggiungi la cartella scripts al path per importare i moduli Turin Wealth
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from tw_config import COLORS, AZIENDE, NUMBER_FORMATS
from tw_utils import set_formula, fmt_header, protect_sheet

# Directory per cache dati e output
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dati_cache")
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


def scarica_o_cache(ticker, period="5y"):
    """Scarica dati da Yahoo Finance, con fallback su cache locale."""
    try:
        df = yf.download(ticker, period=period, progress=False)
        if len(df) > 0:
            return df
    except Exception:
        pass
    # Fallback: leggi dalla cache locale
    nome_file = ticker.replace("^", "").replace(".", "_") + ".csv"
    path_cache = os.path.join(CACHE_DIR, nome_file)
    if os.path.exists(path_cache):
        print(f"Uso cache locale per {ticker}")
        return pd.read_csv(path_cache, index_col=0, parse_dates=True)
    raise FileNotFoundError(f"Nessun dato per {ticker}")


# ------------------------------------------------------------------------------
# SETUP: Carica dati azioni (riprende dal Blocco 1)
# ------------------------------------------------------------------------------
print("=" * 60)
print("BLOCCO 2: Dal DataFrame a Excel")
print("=" * 60)
print()
print("Caricamento dati in corso...")

tickers = [az["ticker"] for az in AZIENDE]
nomi = [az["nome"] for az in AZIENDE]

dati_list = []
for t in tickers:
    df = scarica_o_cache(t)
    dati_list.append(df["Close"].rename(t))

prezzi = pd.concat(dati_list, axis=1)
prezzi.columns = nomi

print(f"Dati caricati: {len(prezzi)} giorni x {len(prezzi.columns)} titoli")
print(f"Periodo: {prezzi.index[0].date()} -> {prezzi.index[-1].date()}")
print()
print(prezzi.tail(3).to_string())
print()

# ------------------------------------------------------------------------------
# Apri Excel (visible=True: lo vediamo mentre il codice gira)
# Gestisce il caso in cui Excel sia gia' aperto
# ------------------------------------------------------------------------------
print("Apertura Excel...")
try:
    app = xw.App(visible=True)
    time.sleep(0.5)  # Lascia inizializzare Excel
except Exception as e:
    print(f"Attenzione: {e}")
    # Se c'e' gia' un'istanza, prova a usarla
    apps = xw.apps
    if apps:
        app = apps.active
    else:
        raise

wb = app.books.add()


# ==============================================================================
# ES06: Da pandas a Excel
#
# Il pattern fondamentale di xlwings:
#   ws.range("A1").value = dataframe          # Scrivi
#   ws.range("A1").options(pd.DataFrame, expand="table").value   # Leggi
#
# Il parametro .options(index=True, header=True) controlla se scrivere
# anche l'indice e le intestazioni delle colonne.
# ==============================================================================
print("=" * 60)
print("ES06 DEMO 1: Scrittura DataFrame in Excel")
print("=" * 60)

ws = wb.sheets[0]
ws.name = "Prezzi"

# Titolo principale
ws.range("A1").value = "QUOTAZIONI STORICHE - Turin Wealth Advisory"
ws.range("A1").font.size = 14
ws.range("A1").font.bold = True

# Scrivi il DataFrame a partire da A3
# options(index=True): include la colonna delle date
# options(header=True): include i nomi delle colonne come prima riga
ultimi_60 = prezzi.tail(60)
ws.range("A3").options(index=True, header=True).value = ultimi_60

print(f"Scritte {len(ultimi_60)} righe x {len(ultimi_60.columns)} colonne")
print(f"Range occupato: A3 -> {chr(65 + len(ultimi_60.columns))}{3 + len(ultimi_60)}")
print()

# ------------------------------------------------------------------------------
# ES06 DEMO 2: Lettura da Excel a pandas
#
# xlwings converte automaticamente:
#   datetime pandas --> date Excel
#   float64 --> numeri Excel
#   NaN --> celle vuote
#   str --> testo Excel
# ------------------------------------------------------------------------------
print("ES06 DEMO 2: Lettura da Excel a pandas")
# expand="table" trova automaticamente i bordi del range
dati_letti = ws.range("A3").options(pd.DataFrame, expand="table").value
print(f"Tipo ritornato: {type(dati_letti)}")
print(f"Shape: {dati_letti.shape}")
print(dati_letti.head(3).to_string())
print()

# ------------------------------------------------------------------------------
# ES06 ESERCIZIO
# ------------------------------------------------------------------------------
print("=" * 60)
print("ES06 ESERCIZIO")
print("=" * 60)
print("""
Obiettivo:
1. Calcola i rendimenti giornalieri dal DataFrame prezzi
   (usa df.pct_change())
2. Aggiungi un nuovo foglio al workbook e chiamalo "Rendimenti"
3. Scrivi il titolo "RENDIMENTI GIORNALIERI - Turin Wealth Advisory" in A1
4. Scrivi i rendimenti a partire dalla cella A3 (con indice e intestazioni)
""")

if MOSTRA_SOLUZIONI:
    # SOLUZIONE ES06
    rendimenti = prezzi.pct_change().dropna()
    ws_rend = wb.sheets.add("Rendimenti")
    ws_rend.range("A1").value = "RENDIMENTI GIORNALIERI - Turin Wealth Advisory"
    ws_rend.range("A1").font.size = 14
    ws_rend.range("A1").font.bold = True
    ws_rend.range("A3").options(index=True, header=True).value = rendimenti
    print(f"[SOLUZIONE] Foglio Rendimenti creato: {len(rendimenti)} righe x {len(rendimenti.columns)} colonne")
    print(f"Periodo: {rendimenti.index[0].date()} -> {rendimenti.index[-1].date()}")
else:
    # Spazio per la soluzione dello studente
    rendimenti = None  # ???  (calcola i rendimenti percentuali giornalieri da prezzi)
    ws_rend = None     # ???  (aggiungi un nuovo foglio al workbook wb)
    # ws_rend.name = ???  (assegna il nome "Rendimenti")
    # Scrivi il titolo in A1
    # Scrivi i rendimenti a partire da A3 (con index e header)
    print(">>> Completa l'esercizio ES06! Imposta MOSTRA_SOLUZIONI = True per vedere la soluzione.")

    # Crea comunque il foglio Rendimenti per permettere agli esercizi successivi di girare
    rendimenti = prezzi.pct_change().dropna()
    ws_rend = wb.sheets.add("Rendimenti")
    ws_rend.range("A1").value = "RENDIMENTI GIORNALIERI - Turin Wealth Advisory"
    ws_rend.range("A1").font.size = 14
    ws_rend.range("A1").font.bold = True
    ws_rend.range("A3").options(index=True, header=True).value = rendimenti

print()
input("Premi Invio per continuare al prossimo esercizio...")


# ==============================================================================
# ES07: Formattazione professionale
#
# La formattazione in xlwings segue la stessa logica di Excel:
#   .color              --> sfondo cella (RGB tuple)
#   .font.color         --> colore testo
#   .font.bold          --> grassetto
#   .api.NumberFormatLocal --> formato numero (versione italiana!)
#
# ATTENZIONE CRITICA: NumberFormatLocal vs number_format
#
#   SBAGLIATO su Excel italiano:
#     rng.number_format = "0.0%"    # Il . viene letto come SEPARATORE MIGLIAIA!
#     rng.number_format = "#,##0"   # La , viene letta come DECIMALE!
#
#   CORRETTO su Excel italiano:
#     rng.api.NumberFormatLocal = "#.##0,00 €"  # . = migliaia, , = decimale
#     rng.api.NumberFormatLocal = "0,00%"       # percentuale con 2 decimali
#     rng.api.NumberFormatLocal = "+0,0%;-0,0%" # con segno
#
# Regola: usare SEMPRE api.NumberFormatLocal con sintassi italiana.
# ==============================================================================
print("=" * 60)
print("ES07 DEMO 1: Formattazione foglio Prezzi")
print("=" * 60)

ws = wb.sheets["Prezzi"]

# Colori dal tema Turin Wealth (da tw_config.py)
HEADER_COLOR = COLORS["header"]       # (44, 62, 80)   - blu profondo
ACCENT       = COLORS["accent"]       # (231, 76, 60)  - corallo energetico
GOLD         = COLORS["gold"]         # (243, 156, 18) - ambra
WHITE        = COLORS["header_text"]  # (255, 255, 255) - bianco
ALT_ROW      = COLORS["table_alt_row"]# (248, 249, 250) - grigio chiaro

print("Colori caricati da tw_config.py:")
print(f"  Header: RGB{HEADER_COLOR}")
print(f"  Accent: RGB{ACCENT}")
print(f"  Gold:   RGB{GOLD}")

# Formatta riga di intestazione (riga 3)
header_range = ws.range("A3").expand("right")
header_range.color = HEADER_COLOR
header_range.font.color = WHITE
header_range.font.bold = True
header_range.font.size = 11

# Formato numeri - ATTENZIONE: Excel italiano!
ultima_riga = ws.range("A3").expand("down").last_cell.row
data_range = ws.range(f"B4:F{ultima_riga}")
data_range.api.NumberFormatLocal = "#.##0,00 €"  # . = migliaia, , = decimale

# Righe alternate per leggibilita'
for i in range(4, ultima_riga + 1):
    if i % 2 == 0:
        ws.range(f"A{i}:F{i}").color = ALT_ROW

ws.autofit()
print(f"Formattazione applicata a {ultima_riga - 3} righe di dati")
print()

# ES07 DEMO 2: Bordi professionali
# I border_id corrispondono alle posizioni: 7=sinistra, 8=destra, 9=top, 10=bottom
print("ES07 DEMO 2: Aggiunta bordi alla tabella")

table = ws.range(f"A3:F{ultima_riga}")
border_color = COLORS["border_gray"]
border_color_int = border_color[0] + border_color[1] * 256 + border_color[2] * 65536

for border_id in [7, 8, 9, 10]:  # Left, Right, Top, Bottom
    border = table.api.Borders(border_id)
    border.LineStyle = 1  # xlContinuous
    border.Weight = 2     # xlThin
    border.Color = border_color_int

print("Bordi applicati alla tabella Prezzi")
print()

# ------------------------------------------------------------------------------
# ES07 ESERCIZIO
# ------------------------------------------------------------------------------
print("=" * 60)
print("ES07 ESERCIZIO")
print("=" * 60)
print("""
Obiettivo:
1. Applica lo stesso header colorato (HEADER_COLOR + WHITE) alla riga 3
   del foglio "Rendimenti"
2. Formato percentuale per i rendimenti: "0,00%" (2 decimali)
   Ricorda: usa api.NumberFormatLocal, NON number_format!
3. Aggiungi formattazione condizionale: celle con rendimento > 0 in verde,
   celle con rendimento < 0 in rosso
   (usa COLORS["correct_bg"] e COLORS["wrong_bg"])
""")

ws_rend = wb.sheets["Rendimenti"]
ultima_riga_rend = ws_rend.range("A3").expand("down").last_cell.row

if MOSTRA_SOLUZIONI:
    # SOLUZIONE ES07
    header_rend = ws_rend.range("A3").expand("right")
    header_rend.color = COLORS["header"]
    header_rend.font.color = COLORS["header_text"]
    header_rend.font.bold = True
    header_rend.font.size = 11

    dati_rend = ws_rend.range(f"B4:F{ultima_riga_rend}")
    dati_rend.api.NumberFormatLocal = "0,00%"

    verde = COLORS["correct_bg"]
    rosso = COLORS["wrong_bg"]
    valori = dati_rend.value
    for r_idx, riga in enumerate(valori):
        for c_idx, val in enumerate(riga):
            if val is not None and isinstance(val, (int, float)):
                cella = ws_rend.range(f"{chr(66 + c_idx)}{4 + r_idx}")
                cella.color = verde if val > 0 else rosso

    ws_rend.autofit()
    print(f"[SOLUZIONE] Formattazione completata: {ultima_riga_rend - 3} righe, colori verde/rosso applicati")
else:
    # 1. Formatta header riga 3
    # ???

    # 2. Formato percentuale per i dati (dalla riga 4)
    # ???

    # 3. Colora le celle: verde se > 0, rosso se < 0
    # ???

    ws_rend.autofit()
    print(">>> Completa l'esercizio ES07! Imposta MOSTRA_SOLUZIONI = True per vedere la soluzione.")

    # Applica comunque la formattazione base per continuare
    ws_rend.range("A3").expand("right").color = COLORS["header"]
    ws_rend.range("A3").expand("right").font.color = COLORS["header_text"]
    ws_rend.range("A3").expand("right").font.bold = True
    ws_rend.range(f"B4:F{ultima_riga_rend}").api.NumberFormatLocal = "0,00%"
    ws_rend.autofit()

print()
input("Premi Invio per continuare al prossimo esercizio...")


# ==============================================================================
# ES08: Formule italiane da Python
#
# Il problema dell'operatore @:
# Excel 365 aggiunge automaticamente @ davanti alle formule legacy:
#   =@MEDIA(B4:F4)   --> quello che vedi (sbagliato!)
#   =MEDIA(B4:F4)    --> quello che volevi (corretto)
#
# La soluzione: set_formula() da tw_utils.py usa Formula2Local invece di
# FormulaLocal, evitando l'aggiunta dell'@.
#
# Sintassi italiana - separatore = punto e virgola (;):
#   IT: =SE(A1>10;"Grande";"Piccolo")
#   EN: =IF(A1>10,"Grande","Piccolo")
#   IT: =CONTA.SE(A:A;">0")
#   EN: =COUNTIF(A:A,">0")
# ==============================================================================
print("=" * 60)
print("ES08 DEMO 1: Il problema dell'@ e la soluzione set_formula()")
print("=" * 60)

ws = wb.sheets["Prezzi"]
ultima_riga = ws.range("A3").expand("down").last_cell.row

# SBAGLIATO - FormulaLocal aggiunge @ in Excel 365
ws.range("G3").value = "Media (con @)"
ws.range("G3").font.bold = True
ws.range("G3").color = COLORS["wrong_bg"]
ws.range("G4").api.FormulaLocal = "=MEDIA(B4:F4)"

time.sleep(1)  # Pausa per vedere il risultato

formula_scritta = ws.range("G4").api.FormulaLocal
print(f"Formula scritta da FormulaLocal: '{formula_scritta}'")
print(f"Ha aggiunto @? {'Si!' if '@' in formula_scritta else 'No'}")

# CORRETTO - set_formula() usa Formula2Local
ws.range("H3").value = "Media (corretta)"
ws.range("H3").font.bold = True
ws.range("H3").color = COLORS["correct_bg"]
set_formula(ws.range("H4"), "=MEDIA(B4:F4)")

time.sleep(1)
formula_corretta = ws.range("H4").api.FormulaLocal
print(f"Formula scritta da set_formula(): '{formula_corretta}'")
print(f"Ha aggiunto @? {'Si!' if '@' in formula_corretta else 'No'}")
print()

# ES08 DEMO 2: Colonne calcolate con formule italiane
# Formula2Local: la API moderna di Excel 365, niente @.
# Regola pratica: usa sempre set_formula() e non preoccuparti dei dettagli interni.
print("ES08 DEMO 2: Aggiunta colonne MEDIA/MIN/MAX con formule italiane")

# Rimuovi le colonne G/H di demo, usa G come Media definitiva
ws.range(f"G3:H{ultima_riga}").clear()

ws.range("G3").value = "Media"
ws.range("H3").value = "Min"
ws.range("I3").value = "Max"

fmt_header(ws.range("G3:I3"))

# NOTA: il separatore e' il PUNTO E VIRGOLA (;) non la virgola!
for row in range(4, ultima_riga + 1):
    set_formula(ws.range(f"G{row}"), f"=MEDIA(B{row}:F{row})")
    set_formula(ws.range(f"H{row}"), f"=MIN(B{row}:F{row})")
    set_formula(ws.range(f"I{row}"), f"=MAX(B{row}:F{row})")

ws.range(f"G4:I{ultima_riga}").api.NumberFormatLocal = "#.##0,00 €"
ws.autofit()
print(f"Aggiunte 3 colonne calcolate (righe 4-{ultima_riga}) con formule MEDIA/MIN/MAX")
print("Esempio formula in G4:", ws.range("G4").api.FormulaLocal)
print()

# ------------------------------------------------------------------------------
# ES08 ESERCIZIO
# ------------------------------------------------------------------------------
print("=" * 60)
print("ES08 ESERCIZIO")
print("=" * 60)
print("""
Obiettivo:
1. Aggiungi una colonna "Positivi" in G che conta quanti rendimenti > 0
   nella riga (quanti titoli sono saliti quel giorno).
   Funzione italiana: CONTA.SE con criterio ">0"
   RICORDA: separatore = punto e virgola (;)
   Struttura: =CONTA.SE(B4:F4;">0")

2. Aggiungi una colonna "Media" in H con la media dei rendimenti della riga.
   Funzione italiana: MEDIA

IMPORTANTE: usa set_formula() per entrambe le colonne!
""")

ws_rend = wb.sheets["Rendimenti"]
ultima_riga_rend = ws_rend.range("A3").expand("down").last_cell.row

ws_rend.range("G3").value = "Positivi"
ws_rend.range("H3").value = "Media"
fmt_header(ws_rend.range("G3:H3"))

if MOSTRA_SOLUZIONI:
    # SOLUZIONE ES08
    for row in range(4, ultima_riga_rend + 1):
        set_formula(ws_rend.range(f"G{row}"), f'=CONTA.SE(B{row}:F{row};">0")')
    for row in range(4, ultima_riga_rend + 1):
        set_formula(ws_rend.range(f"H{row}"), f"=MEDIA(B{row}:F{row})")

    ws_rend.range(f"G4:G{ultima_riga_rend}").api.NumberFormatLocal = "0"
    ws_rend.range(f"H4:H{ultima_riga_rend}").api.NumberFormatLocal = "0,00%"
    ws_rend.autofit()

    formula_g4 = ws_rend.range("G4").api.FormulaLocal
    formula_h4 = ws_rend.range("H4").api.FormulaLocal
    print(f"[SOLUZIONE] Formula G4 (Positivi): {formula_g4}")
    print(f"[SOLUZIONE] Formula H4 (Media):    {formula_h4}")
    ha_chiocciola = "@" in formula_g4 or "@" in formula_h4
    print(f"Presenza @: {'SI (problema!)' if ha_chiocciola else 'NO (corretto!)'}")
else:
    # 1. Colonna Positivi: CONTA.SE con criterio ">0"
    for row in range(4, ultima_riga_rend + 1):
        set_formula(ws_rend.range(f"G{row}"),
            # ??? formula CONTA.SE - separatore italiano = ;
        )

    # 2. Colonna Media: MEDIA dei rendimenti
    for row in range(4, ultima_riga_rend + 1):
        set_formula(ws_rend.range(f"H{row}"),
            # ??? formula MEDIA
        )

    ws_rend.range(f"G4:G{ultima_riga_rend}").api.NumberFormatLocal = "0"
    ws_rend.range(f"H4:H{ultima_riga_rend}").api.NumberFormatLocal = "0,00%"
    ws_rend.autofit()
    print(">>> Completa l'esercizio ES08! Imposta MOSTRA_SOLUZIONI = True per vedere la soluzione.")

    # Applica comunque le formule per continuare
    for row in range(4, ultima_riga_rend + 1):
        set_formula(ws_rend.range(f"G{row}"), f'=CONTA.SE(B{row}:F{row};">0")')
        set_formula(ws_rend.range(f"H{row}"), f"=MEDIA(B{row}:F{row})")
    ws_rend.range(f"G4:G{ultima_riga_rend}").api.NumberFormatLocal = "0"
    ws_rend.range(f"H4:H{ultima_riga_rend}").api.NumberFormatLocal = "0,00%"
    ws_rend.autofit()

print()
input("Premi Invio per continuare al prossimo esercizio...")


# ==============================================================================
# ES09: Il grafico che parla
#
# xlwings permette di creare grafici completi direttamente da Python,
# usando la stessa API di Excel VBA.
#
# Tipi di grafico:
#   chart.chart_type = "line"    --> Linee (andamento nel tempo)
#   chart.chart_type = "bar"     --> Barre orizzontali (confronto)
#   chart.chart_type = "column"  --> Istogramma verticale
#   chart.chart_type = "pie"     --> Torta (composizione)
#
# L'accessor api[1]:
#   chart.api[0] --> ChartObject (il contenitore/cornice nel foglio)
#   chart.api[1] --> Chart (il grafico vero, con assi, serie, titoli)
# Per modificare titoli, serie e colori si usa sempre chart.api[1].
# Questa e' la stessa API usata in VBA con ActiveChart.
# ==============================================================================
print("=" * 60)
print("ES09 DEMO 1: Grafico a linee - Andamento Prezzi")
print("=" * 60)

ws = wb.sheets["Prezzi"]
ultima_riga = ws.range("A3").expand("down").last_cell.row

# Posizione del grafico: sotto i dati
top_grafico = ws.range(f"A{ultima_riga + 3}").top

chart = ws.charts.add(
    left=ws.range("A1").left,
    top=top_grafico,
    width=650,
    height=370
)
chart.chart_type = "line"
chart.set_source_data(ws.range(f"A3:F{ultima_riga}"))

chart.api[1].HasTitle = True
chart.api[1].ChartTitle.Text = "Andamento Prezzi - 5 Titoli Turin Wealth"
chart.api[1].ChartTitle.Font.Size = 13
chart.api[1].ChartTitle.Font.Bold = True

# Personalizza colori delle serie
colori_grafico = [
    COLORS["chart_primary"],
    COLORS["chart_secondary"],
    COLORS["chart_accent1"],
    COLORS["chart_accent2"],
    COLORS["chart_accent3"],
]

for i, colore in enumerate(colori_grafico):
    try:
        serie = chart.api[1].SeriesCollection(i + 1)  # 1-indexed in Excel
        colore_int = colore[0] + colore[1] * 256 + colore[2] * 65536
        serie.Format.Line.ForeColor.RGB = colore_int
        serie.Format.Line.Weight = 2
    except Exception:
        pass

print(f"Grafico linee creato sotto i dati (riga {ultima_riga + 3})")
print()

# ES09 DEMO 2: Grafico a torta - Distribuzione per ultimo prezzo
print("ES09 DEMO 2: Grafico a torta - Allocazione")

ws_alloc = wb.sheets.add("Allocazione")

nomi_az = list(prezzi.columns)
ultimi_prezzi = prezzi.iloc[-1].values

ws_alloc.range("A1").value = "Azienda"
ws_alloc.range("B1").value = "Ultimo Prezzo"
ws_alloc.range("A1:B1").font.bold = True
ws_alloc.range("A1:B1").color = COLORS["header"]
ws_alloc.range("A1:B1").font.color = COLORS["header_text"]

for i, (nome, prezzo) in enumerate(zip(nomi_az, ultimi_prezzi)):
    ws_alloc.range(f"A{i+2}").value = nome
    ws_alloc.range(f"B{i+2}").value = round(prezzo, 2)

ws_alloc.range("B2:B6").api.NumberFormatLocal = "#.##0,00 €"
ws_alloc.autofit()

chart_pie = ws_alloc.charts.add(left=220, top=10, width=420, height=320)
chart_pie.chart_type = "pie"
chart_pie.set_source_data(ws_alloc.range(f"A1:B{len(nomi_az)+1}"))

chart_pie.api[1].HasTitle = True
chart_pie.api[1].ChartTitle.Text = "Distribuzione per Titolo (ultimo prezzo)"
chart_pie.api[1].ChartTitle.Font.Size = 12

try:
    chart_pie.api[1].SeriesCollection(1).DataLabels().ShowPercentage = True
    chart_pie.api[1].SeriesCollection(1).DataLabels().ShowValue = False
except Exception:
    pass

print("Grafico a torta creato nel foglio Allocazione")
print("Titoli:", nomi_az)
print("Ultimi prezzi:", [f"{p:.2f}" for p in ultimi_prezzi])
print()

# ------------------------------------------------------------------------------
# ES09 ESERCIZIO
# ------------------------------------------------------------------------------
print("=" * 60)
print("ES09 ESERCIZIO")
print("=" * 60)
print("""
Obiettivo:
1. Crea un nuovo foglio chiamato "Confronto"
2. Calcola il rendimento annualizzato per ogni titolo:
     rendimento_totale = (ultimo_prezzo / primo_prezzo) - 1
     anni = len(prezzi) / 252   (252 giorni di borsa per anno)
     rendimento_annualizzato = (1 + rendimento_totale) ** (1/anni) - 1
3. Scrivi nomi e rendimenti annualizzati nel foglio
4. Crea un grafico a barre (chart_type = "bar")
   con titolo "Rendimento Annualizzato - Turin Wealth Portfolio"
""")

if MOSTRA_SOLUZIONI:
    # SOLUZIONE ES09
    ws_conf = wb.sheets.add("Confronto")

    anni = len(prezzi) / 252
    nomi_az = list(prezzi.columns)
    rend_annualizzati = []

    for nome in nomi_az:
        serie = prezzi[nome].dropna()
        primo_prezzo = serie.iloc[0]
        ultimo_prezzo = serie.iloc[-1]
        rend_totale = (ultimo_prezzo / primo_prezzo) - 1
        rend_ann = (1 + rend_totale) ** (1 / anni) - 1
        rend_annualizzati.append(rend_ann)

    ws_conf.range("A1").value = "Titolo"
    ws_conf.range("B1").value = "Rendimento Annualizzato"
    ws_conf.range("A1:B1").font.bold = True
    ws_conf.range("A1:B1").color = COLORS["header"]
    ws_conf.range("A1:B1").font.color = COLORS["header_text"]

    for i, (nome, rend) in enumerate(zip(nomi_az, rend_annualizzati)):
        ws_conf.range(f"A{i+2}").value = nome
        ws_conf.range(f"B{i+2}").value = rend

    ws_conf.range(f"B2:B{len(nomi_az)+1}").api.NumberFormatLocal = "0,0%"
    ws_conf.autofit()

    chart_bar = ws_conf.charts.add(left=220, top=10, width=450, height=300)
    chart_bar.chart_type = "bar"
    chart_bar.set_source_data(ws_conf.range(f"A1:B{len(nomi_az)+1}"))

    chart_bar.api[1].HasTitle = True
    chart_bar.api[1].ChartTitle.Text = "Rendimento Annualizzato - Turin Wealth Portfolio"
    chart_bar.api[1].ChartTitle.Font.Size = 12

    try:
        serie_bar = chart_bar.api[1].SeriesCollection(1)
        for i, rend in enumerate(rend_annualizzati):
            colore = COLORS["chart_accent1"] if rend > 0 else COLORS["chart_accent3"]
            colore_int = colore[0] + colore[1] * 256 + colore[2] * 65536
            serie_bar.Points(i + 1).Format.Fill.ForeColor.RGB = colore_int
    except Exception:
        pass

    print("[SOLUZIONE] Grafico confronto creato!")
    for nome, rend in zip(nomi_az, rend_annualizzati):
        segno = "+" if rend > 0 else ""
        print(f"  {nome}: {segno}{rend:.1%}")
else:
    # 1. Crea nuovo foglio
    ws_conf = None  # ???
    # ws_conf.name = ???

    # 2. Calcola i rendimenti annualizzati
    anni = len(prezzi) / 252
    nomi_az = list(prezzi.columns)
    rend_annualizzati = []
    for nome in nomi_az:
        primo_prezzo = None  # ???
        ultimo_prezzo = None  # ???
        rend_totale = None  # ???
        rend_ann = None  # ???
        rend_annualizzati.append(rend_ann)

    # 3. Scrivi i dati nel foglio
    # ws_conf.range("A1").value = "Titolo"
    # ???

    # 4. Crea grafico a barre
    # chart_bar = ???
    # chart_bar.chart_type = "bar"
    # ???

    print(">>> Completa l'esercizio ES09! Imposta MOSTRA_SOLUZIONI = True per vedere la soluzione.")

    # Crea comunque il foglio per permettere al salvataggio di procedere
    ws_conf = wb.sheets.add("Confronto")
    anni = len(prezzi) / 252
    nomi_az = list(prezzi.columns)
    rend_annualizzati = []
    for nome in nomi_az:
        serie = prezzi[nome].dropna()
        primo_prezzo = serie.iloc[0]
        ultimo_prezzo = serie.iloc[-1]
        rend_totale = (ultimo_prezzo / primo_prezzo) - 1
        rend_ann = (1 + rend_totale) ** (1 / anni) - 1
        rend_annualizzati.append(rend_ann)
    ws_conf.range("A1").value = "Titolo"
    ws_conf.range("B1").value = "Rendimento Annualizzato"
    ws_conf.range("A1:B1").font.bold = True
    ws_conf.range("A1:B1").color = COLORS["header"]
    ws_conf.range("A1:B1").font.color = COLORS["header_text"]
    for i, (nome, rend) in enumerate(zip(nomi_az, rend_annualizzati)):
        ws_conf.range(f"A{i+2}").value = nome
        ws_conf.range(f"B{i+2}").value = rend
    ws_conf.range(f"B2:B{len(nomi_az)+1}").api.NumberFormatLocal = "0,0%"
    ws_conf.autofit()

print()
input("Premi Invio per salvare e concludere...")


# ==============================================================================
# SALVATAGGIO
# ==============================================================================
print("=" * 60)
print("SALVATAGGIO")
print("=" * 60)

output_path = os.path.join(OUTPUT_DIR, "report_blocco2.xlsx")
wb.save(output_path)
print(f"File salvato: {output_path}")
print("Fogli presenti nel workbook:")
for s in wb.sheets:
    print(f"  - {s.name}")
print()
print("NON chiudiamo Excel: il file verra' usato nel Blocco 3!")
# wb.close()
# app.quit()


# ==============================================================================
# RIEPILOGO BLOCCO 2
#
# | ES06 | Esportare DataFrame in Excel  | ws.range().options(index=True, header=True).value = df |
# | ES06 | Leggere da Excel in pandas    | .options(pd.DataFrame, expand="table").value           |
# | ES07 | Formattare con colori TW      | .color, .font.color, .font.bold                        |
# | ES07 | Formato numeri italiano       | .api.NumberFormatLocal = "#.##0,00 €"                  |
# | ES08 | Scrivere formule senza @      | set_formula(rng, "=MEDIA(B4:F4)")                      |
# | ES08 | Sintassi italiana Excel       | Separatore ; non ,                                     |
# | ES09 | Creare grafici da Python      | ws.charts.add(...), chart.chart_type                   |
# | ES09 | Personalizzare serie grafici  | chart.api[1].SeriesCollection(i).Format.Line           |
#
# Regole da ricordare:
# 1. NumberFormatLocal con sintassi italiana (. = migliaia, , = decimale)
# 2. set_formula() invece di .api.FormulaLocal per evitare @
# 3. Punto e virgola (;) come separatore nelle formule italiane
# 4. chart.api[1] per accedere al Chart nativo e personalizzarlo
# 5. COLORS da tw_config.py per coerenza visiva con tutto il progetto
# ==============================================================================
print("=" * 60)
print("Blocco 2 completato!")
print("Prossimo: Blocco 3 - Report multi-foglio e protezione")
print("=" * 60)
