# -*- coding: utf-8 -*-
"""
Blocco 1: Dal Web al DataFrame
Python per Excel — Turin Wealth Advisory

Scarica e analizza dati storici di 5 titoli azionari con yfinance e pandas.

Esegui dalla cartella lezione/:
    python 01_dal_web_al_dataframe.py
"""

# ============================================================
# SETUP: Librerie e configurazione
# ============================================================

import yfinance as yf
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys
import os
import warnings
warnings.filterwarnings('ignore')

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))

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

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dati_cache')
os.makedirs(CACHE_DIR, exist_ok=True)


def scarica_o_cache(ticker, period='5y'):
    """Scarica dati da Yahoo Finance; se fallisce, carica dalla cache locale."""
    try:
        df = yf.download(ticker, period=period, progress=False, auto_adjust=True)
        if len(df) > 0:
            # Appiattisci MultiIndex colonne (yfinance v0.2+)
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            nome_file = ticker.replace('^', '').replace('.', '_') + '.csv'
            df.to_csv(os.path.join(CACHE_DIR, nome_file))
            return df
    except Exception as e:
        print(f'  Download fallito per {ticker}: {e}')

    nome_file = ticker.replace('^', '').replace('.', '_') + '.csv'
    path_cache = os.path.join(CACHE_DIR, nome_file)
    if os.path.exists(path_cache):
        print(f'  Uso cache locale per {ticker}')
        return pd.read_csv(path_cache, index_col=0, parse_dates=True)

    raise FileNotFoundError(
        f'Nessun dato disponibile per {ticker}. '
        f'Connettiti a Internet o aggiungi {nome_file} nella cartella dati_cache.'
    )


print(f'\nSetup completato! Aziende: {len(AZIENDE)}, Cache: {CACHE_DIR}')


# ============================================================
# ES01: Scarica Terna e Ferrari
# ============================================================

print('\n' + '='*60)
print('ES01: Il primo ticker')
print('='*60)

print('\nScarico dati Terna (TRN.MI)...')
dati_terna = scarica_o_cache('TRN.MI', period='5y')

print(f'Dimensioni : {dati_terna.shape}  (righe x colonne)')
print(f'Prima data : {dati_terna.index[0].date()}')
print(f'Ultima data: {dati_terna.index[-1].date()}')
print(f'Colonne: {dati_terna.columns.tolist()}')
print('\nPrime 5 righe:')
print(dati_terna.head().to_string())

dati_ferrari = scarica_o_cache('RACE.MI', period='5y')
print(f'\nUltime 10 righe — Ferrari (RACE.MI):')
print(dati_ferrari.tail(10).to_string())
print(f'\nGiorni di borsa: {len(dati_ferrari)}')
print(f"Periodo: {dati_ferrari.index[0].date()} -> {dati_ferrari.index[-1].date()}")
print(f"\nConfronto ultima chiusura:")
print(f"  Terna  : {dati_terna['Close'].iloc[-1]:.2f} EUR")
print(f"  Ferrari: {dati_ferrari['Close'].iloc[-1]:.2f} EUR")


# ============================================================
# ES02: Cinque aziende — download multiplo
# ============================================================

print('\n' + '='*60)
print('ES02: Cinque aziende')
print('='*60)

tickers = [az['ticker'] for az in AZIENDE]
nomi    = [az['nome']   for az in AZIENDE]

print(f'Ticker: {tickers}')
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
elif 'prezzi_close' not in dir():
    prezzi_close = dati_multi
    prezzi_close.columns = nomi

prezzi_close = prezzi_close.dropna(how='all')
print(f'\nTabella prezzi: {prezzi_close.shape}')
print(prezzi_close.head().to_string())

# Analisi dataset
print(f'\nDimensioni: {prezzi_close.shape[0]} righe x {prezzi_close.shape[1]} colonne')

nan_per_titolo = prezzi_close.isna().sum()
print('\nValori mancanti:')
for titolo, nan_count in nan_per_titolo.items():
    print(f'  {titolo:12s}: {nan_count} NaN')

prezzi_completi = prezzi_close.dropna(how='any')
print(f'\nRange comune: {prezzi_completi.index[0].date()} -> {prezzi_completi.index[-1].date()} ({len(prezzi_completi)} giorni)')

print('\nPrezzi piu recenti:')
for titolo, prezzo in prezzi_close.iloc[-1].items():
    valuta = 'USD' if titolo in ['Microsoft', 'Alphabet'] else 'EUR'
    print(f'  {titolo:12s}: {prezzo:8.2f} {valuta}')


# ============================================================
# ES03: Benchmark e normalizzazione a base 100
# ============================================================

print('\n' + '='*60)
print('ES03: Il benchmark')
print('='*60)

print('Scarico benchmark...')
frames_bench = {}
for ticker, nome in zip(['^GSPC', '^STOXX50E'], ['S&P 500', 'Euro Stoxx 50']):
    try:
        df_b = scarica_o_cache(ticker, period='5y')
        frames_bench[nome] = df_b['Close']
    except Exception as e:
        print(f'  Errore {ticker}: {e}')

bench = pd.DataFrame(frames_bench).dropna(how='all')
print(f'Benchmark caricati: {bench.columns.tolist()}')

# Benchmark normalizzati
bench_allineato = bench.dropna(how='any')
bench_norm = bench_allineato / bench_allineato.iloc[0] * 100

fig, ax = plt.subplots(figsize=(11, 5))
for col, colore in zip(bench_norm.columns, ['#E74C3C', '#3498DB']):
    ax.plot(bench_norm.index, bench_norm[col], label=col, color=colore, linewidth=2)
ax.axhline(y=100, color='gray', linestyle='--', linewidth=1, alpha=0.7)
ax.set_title('Benchmark normalizzati (base 100)', fontsize=14, fontweight='bold')
ax.set_ylabel('Performance (base 100)')
ax.legend()
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.show()

print('\nPerformance totale:')
for nome, perf in (bench_norm.iloc[-1] - 100).items():
    print(f'  {nome:15s}: {perf:+.1f}%')

# Titoli + benchmark insieme
tutto = pd.concat([prezzi_close, bench], axis=1).dropna(how='any')
tutto_norm = tutto / tutto.iloc[0] * 100

fig, ax = plt.subplots(figsize=(13, 6))
palette_titoli = ['#2C3E50', '#E74C3C', '#3498DB', '#27AE60', '#9B59B6']
palette_bench  = ['#F39C12', '#E67E22']

for col, colore in zip(prezzi_close.columns, palette_titoli):
    if col in tutto_norm.columns:
        ax.plot(tutto_norm.index, tutto_norm[col], label=col, color=colore, linewidth=1.5, alpha=0.85)
for col, colore in zip(bench.columns, palette_bench):
    if col in tutto_norm.columns:
        ax.plot(tutto_norm.index, tutto_norm[col], label=col, color=colore, linewidth=2.5, linestyle='--', alpha=0.9)
ax.axhline(y=100, color='gray', linestyle=':', linewidth=1)
ax.set_title('Titoli vs Benchmark (base 100)', fontsize=14, fontweight='bold')
ax.set_ylabel('Performance (base 100)')
ax.legend(loc='upper left', fontsize=9)
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.show()

perf_totale = (tutto_norm.iloc[-1] - 100).sort_values(ascending=False)
print('\nClassifica performance:')
for i, (nome, perf) in enumerate(perf_totale.items(), 1):
    tipo = 'BENCH' if nome in bench.columns else 'Titolo'
    print(f'  {i}. {nome:15s} ({tipo:6s}): {perf:+.1f}%')


# ============================================================
# ES04: Rendimenti e statistiche
# ============================================================

print('\n' + '='*60)
print('ES04: Rendimenti e statistiche')
print('='*60)

prezzi = prezzi_close.dropna(how='any')
rendimenti = prezzi.pct_change().dropna()
print(f'Rendimenti giornalieri: {rendimenti.shape[0]} osservazioni')

GIORNI_ANNO = 252
rend_ann = rendimenti.mean() * GIORNI_ANNO
vol_ann  = rendimenti.std() * np.sqrt(GIORNI_ANNO)
sharpe   = rend_ann / vol_ann

stats = pd.DataFrame({
    'Rendimento ann.': rend_ann,
    'Volatilita ann.': vol_ann,
    'Sharpe Ratio':    sharpe,
}).round(4).sort_values('Sharpe Ratio', ascending=False)

print(f'\n{"Titolo":12s}  {"Rendimento":>12s}  {"Volatilita":>12s}  {"Sharpe":>8s}')
print('-' * 52)
for titolo, row in stats.iterrows():
    print(f'{titolo:12s}  {row["Rendimento ann."]:>+11.1%}  {row["Volatilita ann."]:>11.1%}  {row["Sharpe Ratio"]:>8.2f}')

# Matrice di correlazione
corr_matrix = rendimenti.corr()
print('\nMatrice di correlazione:')
print(corr_matrix.round(3).to_string())

corr_lower = corr_matrix.where(np.tril(np.ones(corr_matrix.shape), k=-1).astype(bool))
min_corr = corr_lower.min().min()
min_col  = corr_lower.min().idxmin()
min_row  = corr_lower[min_col].idxmin()
max_corr = corr_lower.max().max()
max_col  = corr_lower.max().idxmax()
max_row  = corr_lower[max_col].idxmax()

print(f'\nCoppia meno correlata: {min_row} & {min_col} (corr = {min_corr:.3f})')
print(f'Coppia piu correlata:  {max_row} & {max_col} (corr = {max_corr:.3f})')

corr_ms_googl = corr_matrix.loc['Microsoft', 'Alphabet']
print(f'Correlazione Microsoft-Alphabet: {corr_ms_googl:.3f}')


# ============================================================
# ES05: Dati reali vs simulati
# ============================================================

print('\n' + '='*60)
print('ES05: Dati reali vs simulati')
print('='*60)

path_dati = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'output', 'DATI_Turin_Wealth.xlsx'
)

try:
    dati_simulati = pd.read_excel(path_dati, sheet_name='QUOTAZIONI_AZIONI', index_col=0)
    # Assicura che l'indice sia datetime (gestisce header "Data" residui)
    dati_simulati.index = pd.to_datetime(dati_simulati.index, errors='coerce')
    dati_simulati = dati_simulati[dati_simulati.index.notna()]
    # Converti colonne a numerico (possono contenere header residui)
    for col in dati_simulati.columns:
        dati_simulati[col] = pd.to_numeric(dati_simulati[col], errors='coerce')
    dati_simulati = dati_simulati.dropna(how='all')
    print(f'Dati simulati caricati: {dati_simulati.shape}')
except FileNotFoundError:
    print('File DATI_Turin_Wealth.xlsx non trovato — genero dati simulati di esempio.')
    np.random.seed(42)
    date_range = pd.date_range(start='2020-01-01', end='2024-12-31', freq='B')
    prezzi_sim = {}
    for ticker, p0 in {'TRN.MI': 6.0, 'RACE.MI': 150.0, 'MSFT': 160.0, 'GOOGL': 68.0, 'MC.PA': 400.0}.items():
        prezzi_sim[ticker] = p0 * np.exp(np.cumsum(np.random.normal(0.0003, 0.015, len(date_range))))
    dati_simulati = pd.DataFrame(prezzi_sim, index=date_range)

def confronta_serie(nome, serie_reale, serie_sim, colore_reale, colore_sim):
    """Confronta una serie reale con una simulata, normalizzando a base 100."""
    sr = serie_reale.dropna()
    ss = serie_sim.dropna()

    # Filtra periodo comune
    data_min = max(sr.index.min(), ss.index.min())
    data_max = min(sr.index.max(), ss.index.max())
    sr = sr[(sr.index >= data_min) & (sr.index <= data_max)]
    ss = ss[(ss.index >= data_min) & (ss.index <= data_max)]

    if len(sr) == 0 or len(ss) == 0:
        print(f'  Nessun periodo comune per {nome}, salto confronto.')
        return

    # Normalizza a base 100
    sr_norm = sr / sr.iloc[0] * 100
    ss_norm = ss / ss.iloc[0] * 100

    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(sr_norm.index, sr_norm.values, color=colore_reale, linewidth=2, label=f'{nome} — Reale')
    ax.plot(ss_norm.index, ss_norm.values, color=colore_sim, linewidth=1.5,
            label=f'{nome} — Simulato', linestyle='--', alpha=0.8)
    ax.axhline(y=100, color='gray', linestyle=':', linewidth=1)
    ax.set_title(f'{nome}: reale vs simulato (base 100)', fontsize=13, fontweight='bold')
    ax.legend()
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.show()

    # Statistiche: confronta su date in comune (reale = giornaliero, sim = mensile)
    date_comuni = sr_norm.index.intersection(ss_norm.index)
    if len(date_comuni) > 5:
        corr = sr_norm[date_comuni].corr(ss_norm[date_comuni])
        mae = np.abs(sr_norm[date_comuni] - ss_norm[date_comuni]).mean()
        print(f'  Correlazione {nome}: {corr:.4f}')
        print(f'  MAE: {mae:.2f} punti (base 100)')
    else:
        # Frequenze diverse (mensile vs giornaliero): resample reale a mensile
        sr_mensile = sr.resample('ME').last()
        ss_mensile = ss.resample('ME').last()
        date_m = sr_mensile.index.intersection(ss_mensile.index)
        if len(date_m) > 3:
            sr_m = sr_mensile[date_m] / sr_mensile[date_m].iloc[0] * 100
            ss_m = ss_mensile[date_m] / ss_mensile[date_m].iloc[0] * 100
            corr = sr_m.corr(ss_m)
            mae = np.abs(sr_m - ss_m).mean()
            print(f'  Correlazione {nome} (mensile): {corr:.4f}')
            print(f'  MAE: {mae:.2f} punti (base 100)')


# Trova colonne simulati e confronta
def trova_colonna_sim(dati, keywords):
    for col in dati.columns:
        if any(kw in str(col).upper() for kw in keywords):
            return col
    return None

col_terna_sim = trova_colonna_sim(dati_simulati, ['TRN', 'TERNA'])
if col_terna_sim is None:
    col_terna_sim = dati_simulati.columns[0]

print(f'\nConfronto Terna ({col_terna_sim}):')
confronta_serie('Terna', prezzi_close['Terna'], dati_simulati[col_terna_sim], '#2C3E50', '#E74C3C')

col_ferrari_sim = trova_colonna_sim(dati_simulati, ['RACE', 'FERRARI'])
if col_ferrari_sim is None and len(dati_simulati.columns) > 1:
    col_ferrari_sim = dati_simulati.columns[1]

if col_ferrari_sim is not None:
    print(f'\nConfronto Ferrari ({col_ferrari_sim}):')
    confronta_serie('Ferrari', prezzi_close['Ferrari'], dati_simulati[col_ferrari_sim], '#E74C3C', '#F39C12')


# ============================================================
# Riepilogo
# ============================================================

print('\n' + '='*60)
print('Blocco 1 completato!')
print('='*60)
print("""
Strumenti utilizzati:
  yfinance.download()    - Scaricare dati dal web
  try/except + cache     - Fallback offline
  .head(), .shape        - Esplorare DataFrame
  yf.download(lista)     - Download multiplo + MultiIndex
  / iloc[0] * 100        - Normalizzazione a base 100
  .pct_change()          - Rendimenti giornalieri
  .std() * sqrt(252)     - Volatilita annualizzata
  .corr()                - Matrice di correlazione
  pd.read_excel()        - Leggere file Excel
""")
