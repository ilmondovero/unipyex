# -*- coding: utf-8 -*-
"""
03_pipeline_completa.py
Pipeline Completa — Turin Wealth Advisory

Versione script standalone del notebook 03_pipeline_completa.ipynb.
Esercizi ES10-ES14: funzioni modulari, protezione, report multi-foglio,
export PDF, confronto simulazione vs realta.

Esegui dalla cartella lezione/:
    python 03_pipeline_completa.py
"""

# Imposta MOSTRA_SOLUZIONI = True per vedere le soluzioni degli esercizi
MOSTRA_SOLUZIONI = False

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

# Aggiungi la cartella scripts al path Python
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))

# Importa configurazioni Turin Wealth
from tw_config import COLORS, AZIENDE, BENCHMARK, NUMBER_FORMATS, CLIENTE

# Importa utilities xlwings
from tw_utils import (
    set_formula, fmt_header, fmt_title, protect_sheet,
    create_workbook, save_and_close, write_table, add_sheet,
    hide_sheet, autofit_all
)

# Cartelle di lavoro
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dati_cache")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)))

# Crea cartella cache se non esiste
os.makedirs(CACHE_DIR, exist_ok=True)


def scarica_o_cache(ticker, period="5y"):
    """
    Scarica dati da Yahoo Finance o usa la cache locale.

    Strategia:
    1. Prova a scaricare da yfinance
    2. Se fallisce (no internet, rate limit), carica da file CSV locale
    3. Se nemmeno la cache esiste, lancia FileNotFoundError
    """
    try:
        df = yf.download(ticker, period=period, progress=False)
        if len(df) > 0:
            nome_file = ticker.replace("^", "").replace(".", "_") + ".csv"
            path_cache = os.path.join(CACHE_DIR, nome_file)
            df.to_csv(path_cache)
            return df
    except Exception:
        pass

    # Fallback: cache locale
    nome_file = ticker.replace("^", "").replace(".", "_") + ".csv"
    path_cache = os.path.join(CACHE_DIR, nome_file)
    if os.path.exists(path_cache):
        print(f"  [cache] Uso dati locali per {ticker}")
        return pd.read_csv(path_cache, index_col=0, parse_dates=True)

    raise FileNotFoundError(
        f"Nessun dato disponibile per {ticker} (ne online ne in cache)"
    )


print("Setup completato.")
print(f"Cache dir: {CACHE_DIR}")
print(f"Output dir: {OUTPUT_DIR}")
print(f"Aziende configurate: {[az['nome'] for az in AZIENDE]}")
print(f"Benchmark configurati: {[b['nome'] for b in BENCHMARK]}")


# ============================================================
# ES10: Lo script completo
# Organizza il codice in funzioni riutilizzabili.
# Principio di responsabilita singola: ogni funzione fa una sola cosa.
# ============================================================

print("\n" + "=" * 60)
print("ES10 — FUNZIONI MODULARI")
print("=" * 60)

# ES10 DEMO: Funzioni modulari


def scarica_dati_aziende():
    """
    Scarica prezzi Close per le 5 aziende Turin Wealth.

    Ritorna: DataFrame con colonne = nomi aziende, indice = date
    """
    dati = {}
    for az in AZIENDE:
        df = scarica_o_cache(az["ticker"])
        # Gestisci MultiIndex colonne (yfinance v0.2+)
        if isinstance(df.columns, pd.MultiIndex):
            dati[az["nome"]] = df["Close"][az["ticker"]]
        else:
            dati[az["nome"]] = df["Close"]
        print(f"  {az['nome']}: {len(df)} giorni di dati")
    return pd.DataFrame(dati)


def calcola_statistiche(prezzi):
    """
    Calcola statistiche di performance per ogni titolo.

    Input:  DataFrame prezzi (date x titoli)
    Output: DataFrame statistiche (titoli x metriche)
    """
    rendimenti = prezzi.pct_change().dropna()

    stats = pd.DataFrame({
        "Ultimo prezzo":      prezzi.iloc[-1],
        "Rend. 1M (%)":       ((prezzi.iloc[-1] / prezzi.iloc[-22]) - 1) * 100,
        "Rend. 1Y (%)":       ((prezzi.iloc[-1] / prezzi.iloc[-252]) - 1) * 100,
        "Volatilita ann. (%)": rendimenti.std() * np.sqrt(252) * 100,
        "Sharpe":             (rendimenti.mean() * 252) / (rendimenti.std() * np.sqrt(252)),
    }).round(2)

    return stats


def crea_report_excel(prezzi, stats, output_path):
    """
    Crea il report Excel completo.
    (Scheletro — completato negli esercizi successivi)
    """
    wb = create_workbook(visible=False)
    # ... implementazione nei prossimi esercizi
    return wb


# Test delle funzioni
print("\nScaricamento dati aziende...")
prezzi = scarica_dati_aziende()

print(f"\nDataFrame prezzi: {prezzi.shape[0]} righe x {prezzi.shape[1]} colonne")
print(f"Periodo: {prezzi.index[0].date()} -> {prezzi.index[-1].date()}")

stats = calcola_statistiche(prezzi)
print("\nStatistiche calcolate:")
print(stats.to_string())

# Pipeline pattern: scarica -> calcola -> crea_report
# Nel pattern if __name__ == "__main__" di uno script standalone:
#
#   if __name__ == "__main__":
#       prezzi = scarica_dati_aziende()
#       stats  = calcola_statistiche(prezzi)
#       crea_report_excel(prezzi, stats, "Report_Q1.xlsx")
#       print("Report generato con successo!")
#
# Nel notebook basta eseguire le celle in ordine.


print("\n--- ES10 ESERCIZIO ---")
print("Scrivi la funzione scarica_benchmark() seguendo lo stesso")
print("pattern di scarica_dati_aziende(), ma iterando su BENCHMARK.")
print()
print("BENCHMARK e' una lista di dizionari con chiavi 'nome' e 'ticker'.")
print("Usa scarica_o_cache() per ogni ticker e ritorna un DataFrame.")

if MOSTRA_SOLUZIONI:
    print("\n--- ES10 SOLUZIONE ---")

    def scarica_benchmark():
        """Scarica i dati degli indici benchmark."""
        dati = {}
        for b in BENCHMARK:
            try:
                df = scarica_o_cache(b["ticker"])
                if isinstance(df.columns, pd.MultiIndex):
                    dati[b["nome"]] = df["Close"][b["ticker"]]
                else:
                    dati[b["nome"]] = df["Close"]
                print(f"  {b['nome']}: {len(df)} giorni di dati")
            except FileNotFoundError as e:
                print(f"  [SKIP] {b['nome']}: {e}")
        return pd.DataFrame(dati)

    print("\nScaricamento benchmark...")
    bench = scarica_benchmark()
    print(f"\nBenchmark disponibili: {list(bench.columns)}")
    print(bench.tail(3).to_string())
else:
    def scarica_benchmark():
        """Scarica i dati degli indici benchmark. DA IMPLEMENTARE."""
        # ???
        # Suggerimento: struttura uguale a scarica_dati_aziende()
        # ma itera su BENCHMARK invece di AZIENDE
        pass


input("\nPremi Invio per continuare...")


# ============================================================
# ES11: Protezione e consegna
# Tre livelli: Locked, FormulaHidden, protect_sheet().
# FormulaHidden funziona SOLO dopo protect_sheet().
# ============================================================

print("\n" + "=" * 60)
print("ES11 — PROTEZIONE E CONSEGNA")
print("=" * 60)

print("""
Gerarchia di protezione in Excel:

  1. Cella  → api.Locked = True/False
               (default: True per tutte le celle)
  2. Cella  → api.FormulaHidden = True
               (nasconde formula nella barra)
  3. Foglio → protect_sheet(ws, password)
               (ATTIVA i flag Locked e FormulaHidden)

IMPORTANTE: Locked e FormulaHidden non fanno NIENTE
finche il foglio non e protetto.
""")

print("--- ES11 DEMO: Protezione foglio ---")
print("Creazione workbook demo separato...")

# Crea workbook demo (SEPARATO dal report principale)
app_demo = xw.App(visible=True)
time.sleep(0.5)
wb_demo = app_demo.books.add()
ws = wb_demo.sheets[0]
ws.name = "Report"

# Intestazione
ws.range("A1").value = "REPORT TRIMESTRALE - FAMIGLIA BIANCHI"
fmt_title(ws.range("A1:D1"))
ws.range("A1:D1").merge()

# Tabella dati
ws.range("A3").value = [["Titolo", "Prezzo", "Variazione %"]]
fmt_header(ws.range("A3:C3"))

ws.range("A4").value = [
    ["Terna",     7.85,   0.032],
    ["Ferrari",   420.50, 0.125],
    ["Microsoft", 425.00, 0.087],
]

# Formula nascosta: calcola variazione media
set_formula(ws.range("D4"), "=MEDIA(C4:C6)")
ws.range("D4").api.FormulaHidden = True  # nasconde nella barra formule

print("Prima della protezione:")
print(f"  FormulaHidden attivo? {ws.range('D4').api.FormulaHidden}")
print("  (clicca su D4 e vedi la formula nella barra sopra)")

# Sblocca celle modificabili dal cliente
ws.range("C4:C6").api.Locked = False
print("\nCelle C4:C6 sbloccate (il cliente potra modificarle)")

# Proteggi il foglio
protect_sheet(ws, "password123")
print("\nFoglio protetto!")
print("  - Celle A1:B6 e D4: bloccate (non modificabili)")
print("  - Celle C4:C6: sbloccate (il cliente puo aggiornare i prezzi)")
print("  - Formula in D4: nascosta nella barra formule")

# Fogli very_hidden
# Visible=0 → xlSheetHidden (appare in Formato > Mostra foglio)
# Visible=2 → xlSheetVeryHidden (invisibile anche li, solo via VBA/Python)
print("\n--- ES11 DEMO 2: Foglio very_hidden ---")
ws_segreto = wb_demo.sheets.add("_CORREZIONE")
ws_segreto.range("A1").value = "Risposte corrette - solo docente!"
ws_segreto.range("A2").value = "Terna rendimento: +3.2%"
ws_segreto.range("A3").value = "Ferrari rendimento: +12.5%"

hide_sheet(ws_segreto, very_hidden=True)

fogli_visibili = [s.name for s in wb_demo.sheets if s.api.Visible == -1]
fogli_nascosti = [s.name for s in wb_demo.sheets if s.api.Visible != -1]

print(f"Fogli visibili in Excel: {fogli_visibili}")
print(f"Fogli nascosti (very hidden): {fogli_nascosti}")
print("Prova in Excel: Home > Formato > Nascondi e mostra > Mostra foglio")
print("Il foglio '_CORREZIONE' non appare nell'elenco!")

print("\n--- ES11 ESERCIZIO ---")
print("Aggiungi un foglio 'Budget' al workbook demo e configuralo:")
print("  1. Titolo in A1: 'BUDGET TRIMESTRALE'")
print("  2. Header in A3:B3: ['Voce', 'Importo']")
print("  3. Dati in A4:B6:")
print("       ['Commissioni advisory', 14000]")
print("       ['Spese di custodia',     2800]")
print("       ['Altri costi',           1200]")
print("  4. Formula SOMMA in B7 con FormulaHidden = True")
print("  5. Sblocca SOLO B4:B6 (importi modificabili)")
print("  6. Proteggi il foglio con password 'turin2024'")

if MOSTRA_SOLUZIONI:
    print("\n--- ES11 SOLUZIONE ---")

    ws_budget = wb_demo.sheets.add("Budget")

    # Titolo
    ws_budget.range("A1").value = "BUDGET TRIMESTRALE"
    fmt_title(ws_budget.range("A1:B1"))
    ws_budget.range("A1:B1").merge()

    # Header
    ws_budget.range("A3").value = [["Voce", "Importo"]]
    fmt_header(ws_budget.range("A3:B3"))

    # Dati
    ws_budget.range("A4").value = [
        ["Commissioni advisory", 14000],
        ["Spese di custodia",     2800],
        ["Altri costi",           1200],
    ]

    # Formula totale con formula nascosta
    ws_budget.range("A7").value = "TOTALE"
    ws_budget.range("A7").font.bold = True
    set_formula(ws_budget.range("B7"), "=SOMMA(B4:B6)")
    ws_budget.range("B7").api.FormulaHidden = True
    ws_budget.range("B7").api.NumberFormatLocal = "#.##0 \u20ac"

    # Sblocca solo le celle importo
    ws_budget.range("B4:B6").api.Locked = False
    ws_budget.range("B4:B6").api.NumberFormatLocal = "#.##0 \u20ac"

    # Proteggi
    protect_sheet(ws_budget, "turin2024")

    print("Foglio Budget creato e protetto!")
    print(f"  Locked A1: {ws_budget.range('A1').api.Locked}  (atteso True)")
    print(f"  Locked B4: {ws_budget.range('B4').api.Locked}  (atteso False)")
    print(f"  FormulaHidden B7: {ws_budget.range('B7').api.FormulaHidden}  (atteso True)")

# Chiudi il workbook demo (SEPARATO dal report principale)
wb_demo.close()
app_demo.quit()
print("\nWorkbook demo chiuso.")

input("\nPremi Invio per continuare...")


# ============================================================
# ES12: Report trimestrale Bianchi
# Architettura multi-foglio: Copertina, Statistiche, Grafici.
# ============================================================

print("\n" + "=" * 60)
print("ES12 — REPORT TRIMESTRALE BIANCHI")
print("=" * 60)

print("""
Architettura del report multi-foglio:

  Foglio       Contenuto                  Destinatario
  ----------   -------------------------  ------------------
  Copertina    Branding, cliente, data    Primo impatto visivo
  Statistiche  Tabella metriche           Roberto e Laura
  Grafici      Andamento prezzi 60gg      Tutti
""")

print("--- ES12 DEMO: Creazione report multi-foglio ---")


def crea_report_bianchi(prezzi, stats):
    """
    Genera il report trimestrale completo per la Famiglia Bianchi.

    Input:
        prezzi: DataFrame (date x titoli) - prezzi storici
        stats:  DataFrame (titoli x metriche) - statistiche calcolate
    Output:
        wb: workbook xlwings aperto (da salvare/chiudere dal chiamante)
    """
    wb = create_workbook(visible=True)
    time.sleep(0.5)

    # --------------------------------------------------------
    # Foglio 1: Copertina
    # --------------------------------------------------------
    ws_cop = wb.sheets[0]
    ws_cop.name = "Copertina"

    ws_cop.range("A:A").column_width = 4
    ws_cop.range("B:B").column_width = 28
    ws_cop.range("C:C").column_width = 20

    # Nome societa
    ws_cop.range("B3").value = "TURIN WEALTH ADVISORY"
    ws_cop.range("B3").font.size = 22
    ws_cop.range("B3").font.bold = True
    ws_cop.range("B3").font.color = COLORS["header"]

    ws_cop.range("B4").value = "Wealth Management & Advisory"
    ws_cop.range("B4").font.size = 13
    ws_cop.range("B4").font.italic = True
    ws_cop.range("B4").font.color = COLORS["subheader"]

    # Linea separatrice
    ws_cop.range("B6:D6").api.Borders(9).Weight = 2
    ws_cop.range("B6:D6").api.Borders(9).Color = sum(
        v * (256 ** i) for i, v in enumerate(reversed(COLORS["accent"]))
    )

    # Titolo report
    ws_cop.range("B8").value = "Report Trimestrale"
    ws_cop.range("B8").font.size = 17
    ws_cop.range("B8").font.bold = True

    trimestre = f"Q{((datetime.now().month - 1) // 3) + 1} {datetime.now().year}"
    ws_cop.range("B9").value = trimestre
    ws_cop.range("B9").font.size = 14
    ws_cop.range("B9").font.color = COLORS["accent"]

    # Dati cliente
    ws_cop.range("B12").value = "Cliente:"
    ws_cop.range("B12").font.bold = True
    ws_cop.range("C12").value = CLIENTE["nome"]

    ws_cop.range("B13").value = "Data:"
    ws_cop.range("B13").font.bold = True
    ws_cop.range("C13").value = datetime.now().strftime("%d/%m/%Y")

    ws_cop.range("B14").value = "Consulente:"
    ws_cop.range("B14").font.bold = True
    ws_cop.range("C14").value = CLIENTE["advisor"]

    ws_cop.range("B15").value = "Patrimonio gestito:"
    ws_cop.range("B15").font.bold = True
    ws_cop.range("C15").value = CLIENTE["patrimonio_totale"]
    ws_cop.range("C15").api.NumberFormatLocal = "\u20ac #.##0"

    # --------------------------------------------------------
    # Foglio 2: Statistiche
    # --------------------------------------------------------
    ws_stats = add_sheet(wb, "Statistiche")

    ws_stats.range("A1").value = "ANALISI TITOLI \u2014 " + trimestre
    fmt_title(ws_stats.range("A1:F1"))
    ws_stats.range("A1:F1").merge()
    ws_stats.row_height = 28

    headers = list(stats.columns)
    data = [[idx] + list(row) for idx, row in stats.iterrows()]
    write_table(ws_stats, (2, 1), ["Titolo"] + headers, data)

    n_rows = len(data)
    ws_stats.range(f"B3:B{2 + n_rows}").api.NumberFormatLocal = "#.##0,00 \u20ac"
    ws_stats.range(f"C3:E{2 + n_rows}").api.NumberFormatLocal = "0,00"
    ws_stats.range(f"F3:F{2 + n_rows}").api.NumberFormatLocal = "0,00"

    # --------------------------------------------------------
    # Foglio 3: Grafici
    # --------------------------------------------------------
    ws_graf = add_sheet(wb, "Grafici")

    ws_graf.range("A1").value = "ANDAMENTO PREZZI \u2014 Ultimi 60 giorni"
    fmt_title(ws_graf.range("A1:F1"))
    ws_graf.range("A1:F1").merge()

    # Normalizza a base 100 per confronto visivo
    ultimi_60 = prezzi.tail(60).copy()
    ultimi_60_norm = (ultimi_60 / ultimi_60.iloc[0]) * 100

    ws_graf.range("A3").options(index=True, header=True).value = ultimi_60_norm
    ws_graf.range("A3").api.NumberFormatLocal = "GG/MM/AAAA"

    last_row = 3 + len(ultimi_60_norm)

    top_pos = ws_graf.range(f"A{last_row + 2}").top
    chart = ws_graf.charts.add(left=10, top=top_pos, width=700, height=380)
    chart.chart_type = "line"
    chart.set_source_data(ws_graf.range(f"A3:F{last_row}"))
    chart.api[1].HasTitle = True
    chart.api[1].ChartTitle.Text = "Performance Comparata (base 100)"
    chart.api[1].HasLegend = True
    chart.api[1].Legend.Position = -4107  # xlLegendPositionBottom

    autofit_all(wb)
    return wb


# Genera il report principale (persiste fino alla cella Save)
print("\nGenerazione report trimestrale...")
report = crea_report_bianchi(prezzi, stats)
print(f"Report creato con {len(report.sheets)} fogli: {[s.name for s in report.sheets]}")

# Struttura consigliata per funzioni molto grandi:
# separare ogni sezione in una funzione privata:
#   _crea_copertina(wb), _crea_foglio_statistiche(wb, stats), ecc.

print("\n--- ES12 ESERCIZIO ---")
print("Aggiungi un foglio 'Correlazione' al report con:")
print("  1. Matrice di correlazione dei rendimenti giornalieri")
print("  2. Heatmap colorata:")
print("     corr > 0.8  → verde scuro   (COLORS['correct_text'])")
print("     corr > 0.5  → verde chiaro  (COLORS['correct_bg'])")
print("     corr > 0.3  → giallo        (COLORS['warning_bg'])")
print("     corr <= 0.3 → rosso chiaro  (COLORS['wrong_bg'])")
print("     diagonale   → grigio        (COLORS['locked_bg'])")

if MOSTRA_SOLUZIONI:
    print("\n--- ES12 SOLUZIONE ---")

    # 1. Calcola matrice correlazione
    rendimenti = prezzi.pct_change().dropna()
    corr = rendimenti.corr().round(2)
    print("Matrice di correlazione:")
    print(corr.to_string())

    # 2. Aggiungi foglio
    ws_corr = add_sheet(report, "Correlazione")

    # 3. Titolo
    ws_corr.range("A1").value = "MATRICE DI CORRELAZIONE \u2014 Rendimenti Giornalieri"
    fmt_title(ws_corr.range("A1:F1"))
    ws_corr.range("A1:F1").merge()

    ws_corr.range("A3").options(index=True, header=True).value = corr
    n_titoli = len(corr.columns)
    fmt_header(ws_corr.range(f"A3:{chr(65 + n_titoli)}3"))

    # 4. Colorazione heatmap
    def corr_color(val):
        if val >= 1.0:
            return COLORS["locked_bg"]
        elif val > 0.8:
            return COLORS["correct_text"]
        elif val > 0.5:
            return COLORS["correct_bg"]
        elif val > 0.3:
            return COLORS["warning_bg"]
        else:
            return COLORS["wrong_bg"]

    for i, titolo1 in enumerate(corr.index):
        for j, titolo2 in enumerate(corr.columns):
            val = corr.loc[titolo1, titolo2]
            riga = 4 + i
            col = 2 + j
            cella = ws_corr.range((riga, col))
            cella.api.NumberFormatLocal = "0,00"
            cella.color = corr_color(val)
            if val > 0.8 and val < 1.0:
                cella.font.color = COLORS["header_text"]

    autofit_all(report)

    fogli = [s.name for s in report.sheets]
    print(f"\nFogli del report: {fogli}")
    print("Foglio Correlazione aggiunto con heatmap!")

input("\nPremi Invio per continuare...")


# ============================================================
# ES13: Esporta PDF
# PageSetup properties, codici header/footer, ExportAsFixedFormat.
# ============================================================

print("\n" + "=" * 60)
print("ES13 — ESPORTA PDF")
print("=" * 60)

print("""
Layout di stampa via PageSetup:

  Proprieta          Valore         Significato
  -----------------  -------------  ---------------------------
  Orientation        2              Landscape
  FitToPagesWide     1              Adatta a 1 pagina larghezza
  FitToPagesTall     1              Adatta a 1 pagina altezza
  CenterHorizontally True           Centra orizzontalmente
  LeftHeader         stringa        Intestazione sinistra
  CenterFooter       "&P di &N"     Pagina corrente / totale

Codici header/footer: &D = data, &P = pagina, &N = totale pagine
""")

print("--- ES13 DEMO: Configura layout e esporta PDF ---")

# Configura il layout di stampa per ogni foglio del report
for ws in report.sheets:
    ps = ws.api.PageSetup
    ps.Orientation         = 2       # xlLandscape
    ps.Zoom                = False
    ps.FitToPagesWide      = 1
    ps.FitToPagesTall      = 1
    ps.CenterHorizontally  = True
    ps.LeftMargin          = 36
    ps.RightMargin         = 36
    ps.TopMargin           = 50
    ps.BottomMargin        = 50
    ps.LeftHeader          = '&"Calibri,Bold"Turin Wealth Advisory'
    ps.RightHeader         = "&D"
    ps.CenterFooter        = "Pagina &P di &N"
    ps.LeftFooter          = "Riservato \u2014 Famiglia Bianchi"

print("Layout di stampa configurato per tutti i fogli.")

# Esporta PDF completo
pdf_path = os.path.join(OUTPUT_DIR, "Report_Bianchi_Completo.pdf")
report.api.ExportAsFixedFormat(
    Type=0,
    Filename=pdf_path,
    Quality=0,
    IncludeDocProperties=True,
    IgnorePrintAreas=False,
    OpenAfterPublish=False
)
print(f"PDF esportato: {pdf_path}")

if os.path.exists(pdf_path):
    size_kb = os.path.getsize(pdf_path) / 1024
    print(f"Dimensione PDF: {size_kb:.1f} KB")
else:
    print("[ERRORE] File PDF non creato!")

print("\n--- ES13 ESERCIZIO ---")
print("Esporta in PDF solo i fogli 'Statistiche' e 'Grafici'")
print("(escludi Copertina e Correlazione).")
print()
print("Strategia:")
print("  1. Nascondi temporaneamente i fogli da escludere (api.Visible = 0)")
print("  2. Esporta con ExportAsFixedFormat")
print("  3. Ripristina la visibilita (api.Visible = -1)")
print("  Usa try/finally per garantire il ripristino anche in caso di errore.")
print("  Nome file: 'Report_Solo_Analisi.pdf'")

if MOSTRA_SOLUZIONI:
    print("\n--- ES13 SOLUZIONE ---")

    fogli_da_escludere = ["Copertina", "Correlazione"]
    pdf_parziale = os.path.join(OUTPUT_DIR, "Report_Solo_Analisi.pdf")

    nascosti = []
    try:
        for ws in report.sheets:
            if ws.name in fogli_da_escludere:
                ws.api.Visible = 0
                nascosti.append(ws.name)
                print(f"  Nascosto temporaneamente: {ws.name}")

        report.api.ExportAsFixedFormat(
            Type=0,
            Filename=pdf_parziale,
            Quality=0,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        print(f"\nPDF parziale esportato: {pdf_parziale}")

    finally:
        # Ripristina visibilita — eseguito SEMPRE anche in caso di errore
        for ws in report.sheets:
            if ws.name in nascosti:
                ws.api.Visible = -1
                print(f"  Ripristinato: {ws.name}")

    if os.path.exists(pdf_parziale):
        size_kb = os.path.getsize(pdf_parziale) / 1024
        print(f"Dimensione: {size_kb:.1f} KB")

input("\nPremi Invio per continuare...")


# ============================================================
# ES14: Il confronto finale
# Simulazione vs realta — dati mensili (random.uniform) vs
# dati reali giornalieri (yfinance).
# ============================================================

print("\n" + "=" * 60)
print("ES14 — IL CONFRONTO FINALE")
print("=" * 60)

print("""
Simulazione vs Realta:

  Caratteristica       Simulazione semplice   Dati reali
  -------------------  ---------------------  --------------------------
  Distribuzione rend.  Normale                Fat tails (code spesse)
  Volatilita           Costante               Cluster (ARCH effects)
  Autocorrelazione     Nulla                  Lieve nelle varianze
  Correlazione         Programmata            Variabile nel tempo
  Frequenza dati       Mensile                Giornaliera
""")

print("--- ES14 DEMO: Confronto volatilita ---")

path_dati = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "output", "DATI_Turin_Wealth.xlsx"
)

try:
    sim = pd.read_excel(path_dati, sheet_name="QUOTAZIONI_AZIONI", index_col=0)
    print("Dati simulati caricati da DATI_Turin_Wealth.xlsx")
    print(f"  Periodo simulato:  {sim.index[0]} -> {sim.index[-1]}")
    print(f"  Frequenza:         mensile ({len(sim)} osservazioni)")
    print(f"  Colonne:           {list(sim.columns)}")

    print(f"\nDati reali (yfinance):")
    print(f"  Periodo reale:     {prezzi.index[0].date()} -> {prezzi.index[-1].date()}")
    print(f"  Frequenza:         giornaliera ({len(prezzi)} osservazioni)")
    print(f"  Colonne:           {list(prezzi.columns)}")

    # Confronto volatilita
    print("\n--- Confronto volatilita annualizzata ---")
    rend_sim  = sim.pct_change().dropna()
    rend_real = prezzi.pct_change().dropna()

    # Mensile -> annualizzato: * sqrt(12)
    vol_sim  = rend_sim.std()  * np.sqrt(12) * 100
    # Giornaliero -> annualizzato: * sqrt(252)
    vol_real = rend_real.std() * np.sqrt(252) * 100

    print(f"  {'Titolo':<15} {'Simul. (%)':>12} {'Reale (%)':>12}")
    for az in AZIENDE:
        nome = az["nome"]
        v_s = vol_sim.get(nome, float("nan"))
        v_r = vol_real.get(nome, float("nan"))
        print(f"  {nome:<15} {v_s:>11.1f}% {v_r:>11.1f}%")

    print("\nNote:")
    print("  - I dati simulati usano random.uniform() con seed fisso")
    print("  - I dati reali mostrano cluster di volatilita (ARCH effects)")
    print("  - La simulazione semplice sottostima i rischi di coda (tail risk)")

except FileNotFoundError:
    print("[INFO] File DATI_Turin_Wealth.xlsx non trovato in output/.")
    print("  Per generarlo: cd scripts && python create_dati.py")
    print("  Procedi con il confronto teorico qui sotto.")
except Exception as e:
    print(f"[ERRORE] {e}")

print("\n--- ES14 RIFLESSIONE GUIDATA ---")
print("Rispondi alle domande qui sotto. Non c'e una risposta giusta o sbagliata:")
print("e un esercizio di pensiero critico.\n")

# Studenti: sostituisci "???" con le tue risposte
riflessioni = {
    "domanda_1": (
        "Perche i dati simulati sono utili per le esercitazioni "
        "anche se non corrispondono alla realta?"
    ),
    "risposta_1": "???",

    "domanda_2": (
        "Quali limitazioni ha il modello random.uniform() "
        "rispetto ai modelli finanziari piu avanzati?"
    ),
    "risposta_2": "???",

    "domanda_3": (
        "Se dovessi migliorare la simulazione, quale modello useresti? "
        "(es. GARCH, random walk con drift, Monte Carlo geometrico?)"
    ),
    "risposta_3": "???",
}

print("Riflessioni ES14 — Simulazione vs Realta")
print("=" * 50)
for num in ["1", "2", "3"]:
    domanda = riflessioni.get(f"domanda_{num}", "")
    risposta = riflessioni.get(f"risposta_{num}", "???")
    print(f"\nDomanda {num}: {domanda}")
    print(f"Risposta:   {risposta}")

input("\nPremi Invio per continuare...")


# ============================================================
# SALVA IL REPORT FINALE
# ============================================================

print("\n" + "=" * 60)
print("SALVATAGGIO REPORT FINALE")
print("=" * 60)

output_path = os.path.join(OUTPUT_DIR, "Report_Trimestrale_Bianchi.xlsx")

try:
    report.save(output_path)
    print(f"Report salvato: {output_path}")
    if os.path.exists(output_path):
        size_kb = os.path.getsize(output_path) / 1024
        print(f"Dimensione file: {size_kb:.1f} KB")
except Exception as e:
    print(f"[ERRORE] Salvataggio fallito: {e}")
    print("Assicurati che il file non sia aperto in un altro programma.")
    raise

# Chiudi Excel
try:
    report.close()
    report.app.quit()
    print("Excel chiuso correttamente.")
except Exception:
    print("[INFO] Excel gia chiuso o istanza non trovata.")

print(f"\nFile generati in {OUTPUT_DIR}:")
for f in os.listdir(OUTPUT_DIR):
    if f.endswith((".xlsx", ".pdf")):
        path_f = os.path.join(OUTPUT_DIR, f)
        print(f"  {f}  ({os.path.getsize(path_f) / 1024:.1f} KB)")


# ============================================================
# RIEPILOGO
# Pipeline pattern, competenze acquisite.
# ============================================================

print("\n" + "=" * 60)
print("RIEPILOGO BLOCCO 3")
print("=" * 60)
print("""
Competenze acquisite:

  ES10  Organizzazione in funzioni     def, parametri, return
  ES11  Protezione fogli e celle       api.Locked, FormulaHidden, protect_sheet()
  ES12  Report multi-foglio            add_sheet(), write_table(), grafici xlwings
  ES13  Export PDF                     PageSetup, ExportAsFixedFormat
  ES14  Pensiero critico               Simulazione vs dati reali

Pattern da ricordare (pipeline standard Turin Wealth):

  prezzi = scarica_dati_aziende()              # 1. Scarica
  stats  = calcola_statistiche(prezzi)         # 2. Elabora
  wb     = crea_report_bianchi(prezzi, stats)  # 3. Crea Excel
  wb.api.ExportAsFixedFormat(0, pdf_path)      # 4. Esporta PDF
  wb.save(xlsx_path)                           # 5. Salva
  wb.close(); wb.app.quit()                    # 6. Chiudi

Prossimo blocco:
  Blocco 4 — Analisi fondamentale: P/E, DCF, confronto multipli.
  Alessandro Bianchi vuole investire EUR 200k in azioni singole.
  Dovrai giustificare ogni scelta con i numeri.

Turin Wealth Advisory — Corso universitario di Excel, Facolta di Economia
""")
