#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Blocco 1: Dal Web al DataFrame
Python per Excel — Turin Wealth Advisory

Scenario: Sei Junior Advisor presso Turin Wealth. Il tuo responsabile
Dott. Marco Ferretti ti chiede di scaricare e analizzare i dati storici
di 5 titoli azionari prima dell'incontro con Alessandro Bianchi (venerdì).

Titoli analizzati:
  TRN.MI  — Terna SpA        (Borsa Italiana, Utilities)
  RACE.MI — Ferrari NV       (Borsa Italiana, Luxury Automotive)
  MSFT    — Microsoft Corp   (NASDAQ, Technology)
  GOOGL   — Alphabet Inc     (NASDAQ, Technology)
  MC.PA   — LVMH             (Euronext Paris, Luxury Goods)

Uso: python 01_dal_web_al_dataframe.py
"""

# ══════════════════════════════════════════════════════════════════
# SETUP: Librerie e configurazione
# ══════════════════════════════════════════════════════════════════

import yfinance as yf
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import sys
import os
import warnings
warnings.filterwarnings('ignore')

# ──────────────────────────────────────────────────────────────────
# Controlla se il tuo docente vuole mostrare le soluzioni
# Cambia a True per vedere le soluzioni automaticamente
# ──────────────────────────────────────────────────────────────────
MOSTRA_SOLUZIONI = False

# Aggiunge la cartella scripts al path per importare tw_config
_script_dir = os.path.dirname(os.path.abspath(__file__))
_root_dir   = os.path.dirname(_script_dir)  # cartella genitore di lezione/
sys.path.insert(0, os.path.join(_root_dir, 'scripts'))

try:
    from tw_config import AZIENDE, BENCHMARK, COLORS
    print('tw_config caricato correttamente.')
except ImportError:
    print('tw_config non trovato — uso valori di fallback.')
    AZIENDE = [
        {'nome': 'Terna',     'ticker': 'TRN.MI'},
        {'nome': 'Ferrari',   'ticker': 'RACE.MI'},
        {'nome': 'Microsoft', 'ticker': 'MSFT'},
        {'nome': 'Alphabet',  'ticker': 'GOOGL'},
        {'nome': 'LVMH',      'ticker': 'MC.PA'},
    ]
    BENCHMARK = [
        {'nome': 'S&P 500',       'ticker': '^GSPC'},
        {'nome': 'Euro Stoxx 50', 'ticker': '^STOXX50E'},
    ]
    COLORS = {}

# ──────────────────────────────────────────────────────────────────
# Directory cache per modalita' offline
# ──────────────────────────────────────────────────────────────────
CACHE_DIR = os.path.join(_script_dir, 'dati_cache')
os.makedirs(CACHE_DIR, exist_ok=True)


def scarica_o_cache(ticker, period='5y'):
    """
    Scarica dati da Yahoo Finance; se il download fallisce,
    carica il file CSV dalla cache locale.

    Parametri
    ---------
    ticker : str   — simbolo del titolo (es. 'MSFT', 'TRN.MI')
    period : str   — periodo storico (es. '5y', '1y', '6mo')

    Come funziona (pattern try/except):
      Internet disponibile  ->  scarica da Yahoo Finance  ->  salva in cache
      Internet non disponibile  ->  legge il CSV dalla cache locale

    Nota su auto_adjust=True: i prezzi vengono gia' aggiustati per split
    azionari e dividendi, producendo serie storiche coerenti.

    Ritorna
    -------
    pd.DataFrame con colonne OHLCV
    """
    try:
        df = yf.download(ticker, period=period, progress=False, auto_adjust=True)
        if len(df) > 0:
            nome_file = ticker.replace('^', '').replace('.', '_') + '.csv'
            df.to_csv(os.path.join(CACHE_DIR, nome_file))
            return df
    except Exception as e:
        print(f'  Download fallito per {ticker}: {e}')

    # Fallback su cache locale
    nome_file = ticker.replace('^', '').replace('.', '_') + '.csv'
    path_cache = os.path.join(CACHE_DIR, nome_file)
    if os.path.exists(path_cache):
        print(f'  Uso cache locale per {ticker}')
        return pd.read_csv(path_cache, index_col=0, parse_dates=True)

    raise FileNotFoundError(
        f'Nessun dato disponibile per {ticker}. '
        f'Connettiti a Internet o aggiungi {nome_file} nella cartella dati_cache.'
    )


if __name__ == '__main__':

    print()
    print('Setup completato!')
    print(f'  - Aziende caricate : {len(AZIENDE)}')
    print(f'  - Cache directory  : {CACHE_DIR}')
    print()


    # ══════════════════════════════════════════════════════════════════
    # ES01: Il primo ticker
# ══════════════════════════════════════════════════════════════════
#
# Ferretti: "Inizia con Terna. E' un'utility italiana, gestisce la
# rete di trasmissione elettrica nazionale. Un titolo stabile,
# difensivo. Scarica i dati degli ultimi 5 anni e dimmi cosa ottieni."
#
# Obiettivo: scaricare dati storici di un titolo e capire
# la struttura di un DataFrame pandas.
# ══════════════════════════════════════════════════════════════════

print('=' * 60)
print('ES01: Il primo ticker')
print('=' * 60)

# ──────────────────────────────────────────────────────────────────
# ES01 — DEMO: Scarica Terna (TRN.MI)
# ──────────────────────────────────────────────────────────────────
print()
print('--- DEMO ---')
print('Scarico dati Terna (TRN.MI)...')
dati_terna = scarica_o_cache('TRN.MI', period='5y')

print()
print('=== STRUTTURA DEL DATAFRAME ===')
print(f'Dimensioni : {dati_terna.shape}  (righe x colonne)')
print(f'Prima data : {dati_terna.index[0].date()}')
print(f'Ultima data: {dati_terna.index[-1].date()}')
print()
print('Colonne disponibili:')
print(dati_terna.columns.tolist())
print()

# Anatomia di un DataFrame:
#
#             Open    High     Low   Close   Volume
# Date
# 2020-01-02  6.123   6.200   6.100   6.180  1234567  <- un giorno di borsa
# 2020-01-03  6.190   6.220   6.150   6.200   987654
#
# Colonne:
#   Open   = Prezzo apertura
#   High   = Massimo giornaliero
#   Low    = Minimo giornaliero
#   Close  = Prezzo chiusura (il piu' usato per analisi storica)
#   Volume = Azioni scambiate
#
# Per quasi tutte le analisi useremo solo Close.
#
# Comandi utili:
#   df.head(n)     prime n righe
#   df.tail(n)     ultime n righe
#   df.shape       (righe, colonne)
#   df.columns     lista colonne
#   df.index       indice (date)
#   df.info()      tipi di dati e valori nulli
#   df.describe()  statistiche descrittive
#   df['Close']    seleziona una colonna

print('=== PRIME 5 RIGHE (df.head()) ===')
print(dati_terna.head().to_string())
print()

# ──────────────────────────────────────────────────────────────────
# ES01 — ESERCIZIO: Scarica Ferrari (RACE.MI)
# ──────────────────────────────────────────────────────────────────
#
# Ferretti: "Bene. Ora prova tu con Ferrari.
# Voglio vedere le ultime 10 righe dei dati."
#
# Istruzioni:
#   1. Usa scarica_o_cache() per scaricare i dati di Ferrari
#      Il ticker di Ferrari su Borsa Italiana e' 'RACE.MI'
#   2. Mostra le ultime 10 righe del DataFrame
#   3. Stampa quanti giorni di borsa hai scaricato

print()
print('--- ESERCIZIO ---')
print()

# 1. Scarica i dati di Ferrari
ticker_ferrari = # ???
dati_ferrari = # ???

# 2. Mostra le ultime 10 righe
# ???

# 3. Stampa il numero di giorni disponibili
# ???

# ──────────────────────────────────────────────────────────────────
# ES01 — SOLUZIONE
# ──────────────────────────────────────────────────────────────────
if MOSTRA_SOLUZIONI or input('\nMostra soluzione ES01? (s/n): ').lower() == 's':
    print()
    print('--- SOLUZIONE ES01 ---')

    # 1. Scarica i dati di Ferrari
    ticker_ferrari = 'RACE.MI'
    dati_ferrari = scarica_o_cache(ticker_ferrari, period='5y')

    # 2. Ultime 10 righe
    print('=== ULTIME 10 RIGHE — Ferrari (RACE.MI) ===')
    print(dati_ferrari.tail(10).to_string())

    # 3. Numero di giorni di borsa
    n_giorni    = len(dati_ferrari)
    prima_data  = dati_ferrari.index[0].date()
    ultima_data = dati_ferrari.index[-1].date()
    print()
    print(f'Giorni di borsa disponibili: {n_giorni}')
    print(f'Periodo: dal {prima_data} al {ultima_data}')

    # Confronto rapido con Terna
    print()
    print('Confronto prezzi ultima chiusura:')
    print(f"  Terna  (TRN.MI) : EUR {dati_terna['Close'].iloc[-1]:.2f}")
    print(f"  Ferrari (RACE.MI): EUR {dati_ferrari['Close'].iloc[-1]:.2f}")
else:
    # Scarica comunque Ferrari (serve negli esercizi successivi)
    try:
        dati_ferrari = scarica_o_cache('RACE.MI', period='5y')
    except Exception:
        dati_ferrari = None


# ══════════════════════════════════════════════════════════════════
# ES02: Cinque aziende
# ══════════════════════════════════════════════════════════════════
#
# Ferretti: "Ora costruisci una tabella pulita con solo i prezzi
# di chiusura di tutti e 5 i titoli."
#
# Obiettivo: scaricare piu' ticker simultaneamente e gestire
# la struttura MultiIndex di pandas.
# ══════════════════════════════════════════════════════════════════

print()
print('=' * 60)
print('ES02: Cinque aziende')
print('=' * 60)

# ──────────────────────────────────────────────────────────────────
# ES02 — DEMO: Scarica tutti e 5 i titoli
# ──────────────────────────────────────────────────────────────────
#
# MultiIndex: quando si scaricano piu' ticker, yfinance restituisce
# colonne a due livelli:
#   Livello 0 (tipo dato): Close | Open | High | Low | Volume
#   Livello 1 (ticker):    MSFT GOOGL ... | MSFT GOOGL ... | ...
#
# Per estrarre solo i prezzi di chiusura:
#   prezzi_close = dati_multi['Close']
#
# Attenzione ai NaN: mercati diversi hanno calendari di festivita'
# diversi. Il 4 luglio NYSE e' chiuso, Borsa Italiana e' aperta.
#   df.dropna(how='all')  rimuove righe dove TUTTI i valori sono NaN
#   df.dropna(how='any')  rimuove righe con ALMENO UN NaN
#   df.fillna(method='ffill')  forward fill: porta avanti l'ultimo prezzo valido

print()
print('--- DEMO ---')

tickers = [az['ticker'] for az in AZIENDE]
nomi    = [az['nome']   for az in AZIENDE]

print(f'Ticker da scaricare: {tickers}')
print('Scarico dati...')

try:
    dati_multi = yf.download(tickers, period='5y', progress=False, auto_adjust=True)
    if len(dati_multi) == 0:
        raise ValueError('Dati vuoti')
except Exception as e:
    print(f'Download multiplo fallito: {e}. Uso cache...')
    frames = {}
    for az in AZIENDE:
        try:
            df = scarica_o_cache(az['ticker'], period='5y')
            frames[az['ticker']] = df['Close']
        except Exception:
            print(f"  Saltato {az['ticker']}")
    dati_multi = pd.DataFrame(frames)
    prezzi_close = dati_multi
    prezzi_close.columns = nomi

if isinstance(dati_multi.columns, pd.MultiIndex):
    prezzi_close = dati_multi['Close'].copy()
    prezzi_close = prezzi_close[tickers]
    prezzi_close.columns = nomi

prezzi_close = prezzi_close.dropna(how='all')

print()
print(f'Dimensioni tabella prezzi: {prezzi_close.shape}')
print()
print('Prime 5 righe:')
print(prezzi_close.head().to_string())
print()

# ──────────────────────────────────────────────────────────────────
# ES02 — ESERCIZIO: Analisi del dataset
# ──────────────────────────────────────────────────────────────────
#
# Ferretti: "Prima di fare qualsiasi analisi, devo capire cosa
# abbiamo. Dimmi: quanti giorni di borsa copriamo? Ci sono dati
# mancanti? Qual e' il range di date comune a tutti i titoli?"
#
# Istruzioni:
#   1. Stampa il numero totale di righe e colonne
#   2. Conta i valori mancanti (NaN) per ogni titolo
#      Suggerimento: usa df.isna().sum()
#   3. Trova la prima e ultima data in cui TUTTI i titoli
#      hanno dati (no NaN su nessuna colonna)
#      Suggerimento: usa df.dropna(how='any') poi .index
#   4. Stampa il prezzo di chiusura piu' recente di ogni titolo

print('--- ESERCIZIO ---')
print()

# 1. Dimensioni del dataset
# ???

# 2. Valori mancanti per titolo
# ???

# 3. Range date con dati completi
prezzi_completi    = # ???
data_inizio_comune = # ???
data_fine_comune   = # ???
print(f'Range comune a tutti i titoli: ??? -> ???')

# 4. Prezzi piu' recenti
# ???

# ──────────────────────────────────────────────────────────────────
# ES02 — SOLUZIONE
# ──────────────────────────────────────────────────────────────────
if MOSTRA_SOLUZIONI or input('\nMostra soluzione ES02? (s/n): ').lower() == 's':
    print()
    print('--- SOLUZIONE ES02 ---')

    # 1. Dimensioni
    print(f'Dimensioni: {prezzi_close.shape[0]} righe x {prezzi_close.shape[1]} colonne')
    print()

    # 2. Valori mancanti per titolo
    nan_per_titolo = prezzi_close.isna().sum()
    print('Valori mancanti per titolo:')
    for titolo, nan_count in nan_per_titolo.items():
        print(f'  {titolo:12s}: {nan_count} NaN')
    print()

    # 3. Range date con dati completi su tutti i titoli
    prezzi_completi    = prezzi_close.dropna(how='any')
    data_inizio_comune = prezzi_completi.index[0].date()
    data_fine_comune   = prezzi_completi.index[-1].date()
    n_giorni_comuni    = len(prezzi_completi)

    print('Range comune a tutti i titoli:')
    print(f'  Da  : {data_inizio_comune}')
    print(f'  A   : {data_fine_comune}')
    print(f'  Giorni di borsa: {n_giorni_comuni}')
    print()

    # 4. Prezzi piu' recenti
    print('Prezzi di chiusura piu\' recenti:')
    ultima_riga = prezzi_close.iloc[-1]
    for titolo, prezzo in ultima_riga.items():
        valuta = 'USD' if titolo in ['Microsoft', 'Alphabet'] else 'EUR'
        print(f'  {titolo:12s}: {prezzo:8.2f} {valuta}')


# ══════════════════════════════════════════════════════════════════
# ES03: Il benchmark
# ══════════════════════════════════════════════════════════════════
#
# Ferretti: "Se Ferrari ha guadagnato il 40% in 5 anni, e' tanto o
# poco? Dipende dal benchmark. Scarica S&P 500 ed Euro Stoxx 50 e
# costruiamo un grafico di confronto normalizzato a base 100."
#
# Obiettivo: normalizzare serie storiche a base 100 per confrontare
# asset con prezzi assoluti molto diversi.
#
# Perche' normalizzare a base 100?
#   Ferrari vale EUR 300, Microsoft USD 400, S&P 500 e' a 4000 punti.
#   Non si confrontano prezzi assoluti. Dopo la normalizzazione,
#   tutti partono da 100 e si leggono come percentuali:
#     Valore 140 = +40% rispetto all'inizio
#     Valore  80 = -20% rispetto all'inizio
#
#   Formula: serie_norm[t] = serie[t] / serie[0] * 100
# ══════════════════════════════════════════════════════════════════

print()
print('=' * 60)
print('ES03: Il benchmark')
print('=' * 60)

# ──────────────────────────────────────────────────────────────────
# ES03 — DEMO: Scarica i benchmark e normalizza
# ──────────────────────────────────────────────────────────────────
print()
print('--- DEMO ---')

bench_tickers = ['^GSPC', '^STOXX50E']
bench_nomi    = ['S&P 500', 'Euro Stoxx 50']

print('Scarico benchmark...')
frames_bench = {}
for ticker, nome in zip(bench_tickers, bench_nomi):
    try:
        df_b = scarica_o_cache(ticker, period='5y')
        frames_bench[nome] = df_b['Close']
    except Exception as e:
        print(f'  Errore {ticker}: {e}')

bench = pd.DataFrame(frames_bench)
bench = bench.dropna(how='all')

print(f'Benchmark caricati: {bench.columns.tolist()}')
print()

# Allinea le date (solo date comuni ai benchmark)
bench_allineato = bench.dropna(how='any')

# Normalizza a base 100
bench_norm = bench_allineato / bench_allineato.iloc[0] * 100

# Grafico
fig, ax = plt.subplots(figsize=(11, 5))
colori_bench = ['#E74C3C', '#3498DB']
for col, colore in zip(bench_norm.columns, colori_bench):
    ax.plot(bench_norm.index, bench_norm[col], label=col, color=colore, linewidth=2)
ax.axhline(y=100, color='gray', linestyle='--', linewidth=1, alpha=0.7)
ax.set_title('Benchmark normalizzati (base 100)', fontsize=14, fontweight='bold')
ax.set_ylabel('Performance (base 100)')
ax.legend()
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.show()

print()
performance_finale = bench_norm.iloc[-1] - 100
print('Performance totale nel periodo:')
for nome, perf in performance_finale.items():
    segno = '+' if perf >= 0 else ''
    print(f'  {nome:15s}: {segno}{perf:.1f}%')
print()

# ──────────────────────────────────────────────────────────────────
# ES03 — ESERCIZIO: Titoli + benchmark sullo stesso grafico
# ──────────────────────────────────────────────────────────────────
#
# Ferretti: "Ora metti INSIEME i 5 titoli e i 2 benchmark.
# Cosi' Alessandro vede subito chi ha battuto il mercato e chi no."
#
# Istruzioni:
#   1. Usa pd.concat() per unire prezzi_close e bench in un
#      unico DataFrame (axis=1, colonne affiancate)
#      Considera solo le date comuni (dropna how='any')
#   2. Normalizza tutto a base 100
#   3. Traccia il grafico con:
#      - Titoli: linee piu' sottili
#      - Benchmark: linee tratteggiate piu' spesse
#      - Linea orizzontale a 100 (base)
#      - Legenda, titolo, griglia
#   4. Stampa la performance totale di ogni titolo ordinata
#      dal migliore al peggiore

print('--- ESERCIZIO ---')
print()

# 1. Unisci prezzi_close e bench
tutto = # ???
tutto = tutto.dropna(how='any')  # solo date con tutti i dati

# 2. Normalizza a base 100
tutto_norm = # ???

# 3. Grafico
fig, ax = plt.subplots(figsize=(12, 6))

# Linee per i 5 titoli
for col in prezzi_close.columns:
    if col in tutto_norm.columns:
        ax.plot(# ???)

# Linee tratteggiate per i benchmark
for col in bench.columns:
    if col in tutto_norm.columns:
        ax.plot(# ???, linestyle='--', linewidth=2)

ax.axhline(y=100, color='gray', linestyle=':', linewidth=1)
ax.set_title(# ???)
ax.set_ylabel('Performance (base 100)')
ax.legend(# ???)
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.show()

# 4. Classifica performance totale
print('\nClassifica performance totale (dal migliore al peggiore):')
# ???

# ──────────────────────────────────────────────────────────────────
# ES03 — SOLUZIONE
# ──────────────────────────────────────────────────────────────────
if MOSTRA_SOLUZIONI or input('\nMostra soluzione ES03? (s/n): ').lower() == 's':
    print()
    print('--- SOLUZIONE ES03 ---')

    # 1. Unisci titoli e benchmark
    tutto = pd.concat([prezzi_close, bench], axis=1)
    tutto = tutto.dropna(how='any')

    print(f'Date comuni a titoli + benchmark: {len(tutto)} giorni di borsa')
    print()

    # 2. Normalizza a base 100
    tutto_norm = tutto / tutto.iloc[0] * 100

    # 3. Grafico
    palette_titoli    = ['#2C3E50', '#E74C3C', '#3498DB', '#27AE60', '#9B59B6']
    palette_benchmark = ['#F39C12', '#E67E22']

    fig, ax = plt.subplots(figsize=(13, 6))

    for col, colore in zip(prezzi_close.columns, palette_titoli):
        if col in tutto_norm.columns:
            ax.plot(tutto_norm.index, tutto_norm[col],
                    label=col, color=colore, linewidth=1.5, alpha=0.85)

    for col, colore in zip(bench.columns, palette_benchmark):
        if col in tutto_norm.columns:
            ax.plot(tutto_norm.index, tutto_norm[col],
                    label=col, color=colore, linewidth=2.5,
                    linestyle='--', alpha=0.9)

    ax.axhline(y=100, color='gray', linestyle=':', linewidth=1, label='Base (inizio periodo)')
    ax.set_title('Titoli vs Benchmark — Performance normalizzata (base 100)',
                 fontsize=14, fontweight='bold')
    ax.set_ylabel('Performance (base 100)')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.show()

    # 4. Classifica performance totale
    perf_totale = (tutto_norm.iloc[-1] - 100).sort_values(ascending=False)
    print('Classifica performance totale (dal migliore al peggiore):')
    for i, (nome, perf) in enumerate(perf_totale.items(), 1):
        tipo  = 'BENCHMARK' if nome in bench.columns else 'Titolo'
        segno = '+' if perf >= 0 else ''
        print(f'  {i}. {nome:15s} ({tipo:9s}): {segno}{perf:.1f}%')


# ══════════════════════════════════════════════════════════════════
# ES04: Rendimenti e statistiche
# ══════════════════════════════════════════════════════════════════
#
# Alessandro: "Qual e' il titolo con il miglior rapporto
# rischio/rendimento? Qual e' la correlazione tra Microsoft e Google?"
#
# Obiettivo: calcolare rendimento, volatilita', Sharpe Ratio
# e matrice di correlazione.
#
# Formule:
#   Rendimento giornaliero   : r[t] = (P[t] - P[t-1]) / P[t-1]
#   Rendimento annualizzato  : R_ann = media_giornaliera * 252
#   Volatilita' annualizzata : sigma_ann = std_giornaliera * sqrt(252)
#   Sharpe Ratio             : SR = R_ann / sigma_ann  (senza risk-free)
#
#   Un anno ha circa 252 giorni di borsa (non 365).
#
#   Sharpe Ratio:
#     < 0   : peggio del tasso privo di rischio
#     0 - 1 : accettabile
#     1 - 2 : buono
#     > 2   : eccellente
# ══════════════════════════════════════════════════════════════════

print()
print('=' * 60)
print('ES04: Rendimenti e statistiche')
print('=' * 60)

# ──────────────────────────────────────────────────────────────────
# ES04 — DEMO: Rendimenti e statistiche annualizzate
# ──────────────────────────────────────────────────────────────────
print()
print('--- DEMO ---')

prezzi = prezzi_close.dropna(how='any')

# pct_change() calcola la variazione percentuale giorno su giorno.
# La prima riga sara' NaN (non c'e' il giorno precedente).
rendimenti = prezzi.pct_change().dropna()

print(f'Rendimenti giornalieri: {rendimenti.shape[0]} osservazioni')
print()

GIORNI_ANNO = 252

rend_annualizzato = rendimenti.mean() * GIORNI_ANNO
vol_annualizzata  = rendimenti.std()  * np.sqrt(GIORNI_ANNO)
sharpe_ratio      = rend_annualizzato / vol_annualizzata

stats = pd.DataFrame({
    'Rendimento ann.': rend_annualizzato,
    'Volatilita ann.' : vol_annualizzata,
    'Sharpe Ratio'    : sharpe_ratio,
}).round(4)

stats_ordinata = stats.sort_values('Sharpe Ratio', ascending=False)

print('=== STATISTICHE ANNUALIZZATE ===')
print()
print(f'{"Titolo":12s}  {"Rendimento":>12s}  {"Volatilita":>12s}  {"Sharpe":>8s}')
print('-' * 52)
for titolo, row in stats_ordinata.iterrows():
    r = row['Rendimento ann.']
    v = row['Volatilita ann.']
    s = row['Sharpe Ratio']
    print(f'{titolo:12s}  {r:>+11.1%}  {v:>11.1%}  {s:>8.2f}')

print()
print('(Sharpe Ratio: maggiore = migliore rapporto rischio/rendimento)')
print()

# ──────────────────────────────────────────────────────────────────
# ES04 — ESERCIZIO: Matrice di correlazione
# ──────────────────────────────────────────────────────────────────
#
# Alessandro: "Microsoft e Google sono entrambe tech USA.
# Ha senso tenerle entrambe in portafoglio?"
#
# Istruzioni:
#   1. Calcola la matrice di correlazione dei rendimenti
#      Suggerimento: usa .corr() sul DataFrame dei rendimenti
#   2. Mostra la matrice (df.round(2))
#   3. Trova la coppia di titoli con correlazione MINORE
#      (esclusa la diagonale che e' sempre 1.0)
#      Suggerimento: usa np.tril() per il triangolo inferiore
#   4. Stampa un'interpretazione: la coppia meno correlata
#      e' quella che offre la migliore diversificazione

print('--- ESERCIZIO ---')
print()

# 1. Matrice di correlazione dei rendimenti giornalieri
corr_matrix = # ???

# 2. Mostra la matrice
print('=== MATRICE DI CORRELAZIONE (rendimenti giornalieri) ===')
# ???

# 3. Trova la coppia meno correlata
corr_lower = corr_matrix.where(
    np.tril(np.ones(corr_matrix.shape), k=-1).astype(bool)
)

min_corr = # ???  (usa .min().min() per trovare il valore minimo)
pos_min  = # ???  (usa idxmin() per trovare la posizione)

# 4. Interpretazione
print()
print(f'Correlazione minima: ??? tra ??? e ???')
print('-> Questa coppia offre la migliore diversificazione.')
print()
print('Correlazione Microsoft-Alphabet:')
# ???

# ──────────────────────────────────────────────────────────────────
# ES04 — SOLUZIONE
# ──────────────────────────────────────────────────────────────────
if MOSTRA_SOLUZIONI or input('\nMostra soluzione ES04? (s/n): ').lower() == 's':
    print()
    print('--- SOLUZIONE ES04 ---')

    # 1. Matrice di correlazione
    corr_matrix = rendimenti.corr()

    # 2. Visualizza
    print('=== MATRICE DI CORRELAZIONE (rendimenti giornalieri) ===')
    print()
    print(corr_matrix.round(3).to_string())
    print()

    # 3. Triangolo inferiore (esclude la diagonale)
    corr_lower = corr_matrix.where(
        np.tril(np.ones(corr_matrix.shape), k=-1).astype(bool)
    )

    # Valore minimo
    min_corr = corr_lower.min().min()
    min_col  = corr_lower.min().idxmin()
    min_row  = corr_lower[min_col].idxmin()

    # Valore massimo
    max_corr = corr_lower.max().max()
    max_col  = corr_lower.max().idxmax()
    max_row  = corr_lower[max_col].idxmax()

    # 4. Interpretazione
    print(f'Coppia MENO correlata (migliore diversificazione):')
    print(f'  {min_row} & {min_col}: corr = {min_corr:.3f}')
    print()
    print(f"Coppia PIU' correlata (scarsa diversificazione):")
    print(f'  {max_row} & {max_col}: corr = {max_corr:.3f}')
    print()

    corr_ms_googl = corr_matrix.loc['Microsoft', 'Alphabet']
    print(f'Correlazione Microsoft-Alphabet: {corr_ms_googl:.3f}')
    print()
    if corr_ms_googl > 0.8:
        print('-> Alta correlazione! Tenerle entrambe riduce poco il rischio.')
        print('-> Suggerimento per Alessandro: considera di scegliere una sola fra le due.')
    elif corr_ms_googl > 0.5:
        print('-> Correlazione moderata. Si possono tenere entrambe con buona diversificazione.')
    else:
        print('-> Bassa correlazione. Ottima diversificazione!')


# ══════════════════════════════════════════════════════════════════
# ES05: Dati reali vs simulati
# ══════════════════════════════════════════════════════════════════
#
# Ferretti: "I nostri file Excel usano dati simulati. Ma quanto
# si discostano dalla realta'? Confronta i prezzi di Terna nel
# file Excel con quelli scaricati da Yahoo Finance."
#
# Obiettivo: leggere un file Excel con pandas e confrontare
# dati simulati con dati reali.
#
# pd.read_excel():
#   df = pd.read_excel(
#       'file.xlsx',
#       sheet_name='NomeFoglio',   # nome foglio (o indice numerico)
#       index_col=0,               # prima colonna = indice
#       parse_dates=True,          # interpreta date automaticamente
#   )
#
# Allineamento date:
#   serie.reindex(altra_serie.index).dropna()   # reindex sull'indice comune
#   df1.join(df2, how='inner')                   # join sulle date comuni
# ══════════════════════════════════════════════════════════════════

print()
print('=' * 60)
print('ES05: Dati reali vs simulati')
print('=' * 60)

# ──────────────────────────────────────────────────────────────────
# ES05 — DEMO: Carica i dati simulati dal file Excel
# ──────────────────────────────────────────────────────────────────
print()
print('--- DEMO ---')

path_dati = os.path.join(_root_dir, 'output', 'DATI_Turin_Wealth.xlsx')

print(f'Cerco il file: {path_dati}')
print(f'File esiste: {os.path.exists(path_dati)}')
print()

try:
    dati_simulati = pd.read_excel(
        path_dati,
        sheet_name='QUOTAZIONI_AZIONI',
        index_col=0,
        parse_dates=True
    )
    print(f'Dati simulati caricati: {dati_simulati.shape}')
    print(f'Colonne: {dati_simulati.columns.tolist()}')
    print()
    print('Prime 5 righe dati simulati:')
    print(dati_simulati.head().to_string())

except FileNotFoundError:
    print('File DATI_Turin_Wealth.xlsx non trovato.')
    print('Esegui prima: cd scripts && python create_dati.py')
    print()
    print('Per questo esercizio, generiamo dati simulati di esempio:')

    np.random.seed(42)
    date_range = pd.date_range(start='2020-01-01', end='2024-12-31', freq='B')
    prezzi_sim = {}
    prezzi_iniziali = {
        'TRN.MI': 6.0, 'RACE.MI': 150.0, 'MSFT': 160.0,
        'GOOGL': 68.0, 'MC.PA': 400.0
    }
    for ticker, p0 in prezzi_iniziali.items():
        rend_giorn = np.random.normal(0.0003, 0.015, len(date_range))
        prezzi_sim[ticker] = p0 * np.exp(np.cumsum(rend_giorn))
    dati_simulati = pd.DataFrame(prezzi_sim, index=date_range)
    print(f'Dati simulati di fallback: {dati_simulati.shape}')
    print(dati_simulati.head().to_string())

print()
print(f'Dati reali   : {prezzi_close.shape}')
print(f'Dati simulati: {dati_simulati.shape}')
print()

# ──────────────────────────────────────────────────────────────────
# ES05 — DEMO 2: Confronto grafico Terna (reale vs simulato)
# ──────────────────────────────────────────────────────────────────
print('--- DEMO 2: Confronto Terna ---')
print()

# Individua la colonna Terna nei dati simulati
col_terna_sim = None
for col in dati_simulati.columns:
    if 'TRN' in str(col).upper() or 'TERNA' in str(col).upper():
        col_terna_sim = col
        break

if col_terna_sim is None:
    col_terna_sim = dati_simulati.columns[0]
    print(f'Colonna Terna non trovata per nome, uso: {col_terna_sim}')

terna_reale = prezzi_close['Terna'].dropna()
terna_sim   = dati_simulati[col_terna_sim].dropna()

date_comuni = terna_reale.index.intersection(terna_sim.index)

if len(date_comuni) == 0:
    print('Nessuna data in comune — normalizzo le serie a base 100 per confronto.')
    data_min = max(terna_reale.index.min(), terna_sim.index.min())
    data_max = min(terna_reale.index.max(), terna_sim.index.max())
    terna_reale = terna_reale[(terna_reale.index >= data_min) & (terna_reale.index <= data_max)]
    terna_sim   = terna_sim  [(terna_sim.index   >= data_min) & (terna_sim.index   <= data_max)]
    terna_reale_norm = terna_reale / terna_reale.iloc[0] * 100
    terna_sim_norm   = terna_sim   / terna_sim.iloc[0]   * 100
    usa_normalizzato = True
else:
    terna_reale_norm = terna_reale[date_comuni]
    terna_sim_norm   = terna_sim[date_comuni]
    usa_normalizzato = False

label_y = 'Performance (base 100)' if usa_normalizzato else 'Prezzo (EUR)'

fig, axes = plt.subplots(2, 1, figsize=(12, 8))

ax1 = axes[0]
ax1.plot(terna_reale_norm.index, terna_reale_norm.values,
         color='#2C3E50', linewidth=1.5, label='Terna — Reale (Yahoo Finance)', alpha=0.9)
ax1.plot(terna_sim_norm.index, terna_sim_norm.values,
         color='#E74C3C', linewidth=1.5, label='Terna — Simulato (Excel)', alpha=0.7,
         linestyle='--')
ax1.set_title('Terna: prezzi reali vs simulati', fontsize=13, fontweight='bold')
ax1.set_ylabel(label_y)
ax1.legend()
ax1.grid(True, alpha=0.3)

ax2 = axes[1]
differenza = terna_reale_norm - terna_sim_norm
if not usa_normalizzato:
    colori_diff = ['#27AE60' if v >= 0 else '#E74C3C' for v in differenza]
    ax2.bar(differenza.index, differenza.values, color=colori_diff, alpha=0.6, width=1)
    ax2.axhline(y=0, color='black', linewidth=0.8)
    ax2.set_title('Differenza (Reale - Simulato)', fontsize=13)
    ax2.set_ylabel('Differenza prezzo (EUR)')
else:
    ax2.plot(differenza.index, differenza.values, color='#9B59B6', linewidth=1)
    ax2.axhline(y=0, color='black', linewidth=0.8, linestyle='--')
    ax2.set_title('Differenza normalizzata (Reale - Simulato)', fontsize=13)
    ax2.set_ylabel('Differenza (punti base 100)')
ax2.grid(True, alpha=0.3)

plt.tight_layout()
plt.show()

if not usa_normalizzato:
    corr_terna = terna_reale_norm.corr(terna_sim_norm)
    print(f'Correlazione Terna reale vs simulato: {corr_terna:.4f}')
    if corr_terna > 0.95:
        print('-> Ottima corrispondenza: i dati simulati replicano bene la realta\'.')
    elif corr_terna > 0.80:
        print('-> Buona corrispondenza: andamento simile, prezzi assoluti diversi.')
    else:
        print('-> Scarsa corrispondenza: i simulati si discostano significativamente.')

print()

# ──────────────────────────────────────────────────────────────────
# ES05 — ESERCIZIO: Confronto per Ferrari
# ──────────────────────────────────────────────────────────────────
#
# Ferretti: "Bene per Terna. Ora fai la stessa analisi
# per Ferrari — e' il titolo che interessa di piu' ad Alessandro."
#
# Istruzioni:
#   1. Individua la colonna Ferrari nei dati simulati
#      (cerca 'RACE' o 'Ferrari' nel nome colonna)
#   2. Allinea le date tra serie reale e simulata
#   3. Normalizza entrambe a base 100
#   4. Traccia un grafico con le due serie
#   5. Calcola la correlazione tra le due serie
#      e stampa un'interpretazione
#
# Bonus: calcola anche il MAE (Mean Absolute Error) in punti base 100

print('--- ESERCIZIO ---')
print()

# 1. Trova la colonna Ferrari nei dati simulati
col_ferrari_sim = None
for col in dati_simulati.columns:
    if # ???
        col_ferrari_sim = col
        break

if col_ferrari_sim is None:
    print('Colonna Ferrari non trovata. Colonne disponibili:', dati_simulati.columns.tolist())
else:
    print(f'Colonna Ferrari nei simulati: {col_ferrari_sim}')

    # 2. Allinea le date
    ferrari_reale = # ???
    ferrari_sim   = # ???

    # Normalizza a base 100
    data_start = # ???  (massimo tra le date di inizio delle due serie)
    data_end   = # ???  (minimo tra le date di fine)

    ferrari_reale_norm = # ???
    ferrari_sim_norm   = # ???

    # 3-4. Grafico
    fig, ax = plt.subplots(figsize=(12, 5))
    # ???
    plt.show()

    # 5. Correlazione e interpretazione
    # ???

    # Bonus: MAE
    # ???

# ──────────────────────────────────────────────────────────────────
# ES05 — SOLUZIONE
# ──────────────────────────────────────────────────────────────────
if MOSTRA_SOLUZIONI or input('\nMostra soluzione ES05? (s/n): ').lower() == 's':
    print()
    print('--- SOLUZIONE ES05 ---')

    # 1. Trova la colonna Ferrari nei dati simulati
    col_ferrari_sim = None
    for col in dati_simulati.columns:
        if 'RACE' in str(col).upper() or 'FERRARI' in str(col).upper():
            col_ferrari_sim = col
            break

    if col_ferrari_sim is None and len(dati_simulati.columns) > 1:
        col_ferrari_sim = dati_simulati.columns[1]
        print(f'Colonna Ferrari non trovata per nome, uso: {col_ferrari_sim}')

    if col_ferrari_sim is None:
        print('Impossibile trovare i dati Ferrari simulati.')
    else:
        print(f'Colonna Ferrari trovata: "{col_ferrari_sim}"')

        # 2. Estrai e allinea temporalmente
        ferrari_reale = prezzi_close['Ferrari'].dropna()
        ferrari_sim   = dati_simulati[col_ferrari_sim].dropna()

        data_start = max(ferrari_reale.index.min(), ferrari_sim.index.min())
        data_end   = min(ferrari_reale.index.max(), ferrari_sim.index.max())

        ferrari_reale_fil = ferrari_reale[
            (ferrari_reale.index >= data_start) & (ferrari_reale.index <= data_end)
        ]
        ferrari_sim_fil = ferrari_sim[
            (ferrari_sim.index >= data_start) & (ferrari_sim.index <= data_end)
        ]

        # 3. Normalizza a base 100
        ferrari_reale_norm = ferrari_reale_fil / ferrari_reale_fil.iloc[0] * 100
        ferrari_sim_norm   = ferrari_sim_fil   / ferrari_sim_fil.iloc[0]   * 100

        print(f'Periodo analizzato: {data_start.date()} -> {data_end.date()}')
        print(f'Osservazioni reali   : {len(ferrari_reale_norm)}')
        print(f'Osservazioni simulate: {len(ferrari_sim_norm)}')
        print()

        # 4. Grafico
        fig, ax = plt.subplots(figsize=(12, 5))
        ax.plot(ferrari_reale_norm.index, ferrari_reale_norm.values,
                color='#E74C3C', linewidth=2, label='Ferrari — Reale (Yahoo Finance)')
        ax.plot(ferrari_sim_norm.index, ferrari_sim_norm.values,
                color='#F39C12', linewidth=1.5, label='Ferrari — Simulato (Excel)',
                linestyle='--', alpha=0.8)
        ax.axhline(y=100, color='gray', linestyle=':', linewidth=1)
        ax.set_title('Ferrari NV: prezzi reali vs simulati (base 100)',
                     fontsize=13, fontweight='bold')
        ax.set_ylabel('Performance (base 100)')
        ax.legend()
        ax.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.show()

        # 5. Correlazione
        date_comuni_ferr = ferrari_reale_norm.index.intersection(ferrari_sim_norm.index)

        if len(date_comuni_ferr) > 10:
            corr_ferr = ferrari_reale_norm[date_comuni_ferr].corr(
                ferrari_sim_norm[date_comuni_ferr]
            )
            print(f'Correlazione Ferrari reale vs simulato: {corr_ferr:.4f}')
            print()
            if corr_ferr > 0.95:
                print('-> Eccellente: i simulati replicano molto bene la dinamica reale.')
            elif corr_ferr > 0.80:
                print('-> Buona: andamento simile ma con differenze nei prezzi assoluti.')
            elif corr_ferr > 0.50:
                print('-> Moderata: qualche somiglianza ma i simulati si discostano.')
            else:
                print('-> Bassa: i simulati non replicano bene la dinamica reale.')

            # Bonus: MAE
            mae = np.abs(
                ferrari_reale_norm[date_comuni_ferr] - ferrari_sim_norm[date_comuni_ferr]
            ).mean()
            print()
            print(f'MAE (errore medio assoluto): {mae:.2f} punti (base 100)')
            print(f'-> In media, i simulati si discostano di {mae:.2f} punti rispetto alla base 100.')
        else:
            print('Troppo poche date in comune per calcolare la correlazione.')


# ══════════════════════════════════════════════════════════════════
# RIEPILOGO BLOCCO 1
# ══════════════════════════════════════════════════════════════════
#
# Concetti appresi in questo blocco:
#
#   Scarica dati dal web        yfinance.download()
#   Gestione errori di rete     Pattern try/except con fallback su cache
#   Navigare un DataFrame       .head(), .tail(), .shape, .info(), .describe()
#   Piu' ticker insieme         yf.download(lista) + gestione MultiIndex
#   Confronto serie diverse     Normalizzazione a base 100
#   Rendimenti                  .pct_change()
#   Rischio/rendimento          Rendimento annualizzato, volatilita', Sharpe Ratio
#   Diversificazione            Matrice di correlazione .corr()
#   Leggere file Excel          pd.read_excel()
#
# Prossimo blocco: Analisi fondamentale — bilanci aziendali,
# KPI finanziari (P/E, ROE, EBITDA margin), scoring card per Alessandro.
#
# Ferretti, soddisfatto, chiude il laptop:
# "Bene. Domani mattina alle 9 mi mandi un riepilogo di quello
# che hai trovato. Alessandro arriva venerdi' — voglio essere preparato."
# ══════════════════════════════════════════════════════════════════

print()
print('=' * 60)
print('Blocco 1 completato!')
print('Pronto per il Blocco 2: Analisi fondamentale.')
print('=' * 60)
