"""
genera_cache.py
---------------
Scarica i dati storici finanziari da Yahoo Finance e li salva come file CSV
nella stessa cartella dati_cache/. Usato come fallback offline per i
notebook didattici del corso.

Utilizzo:
    python genera_cache.py

Requisiti:
    pip install yfinance pandas
"""

import yfinance as yf
import pandas as pd
from pathlib import Path

# ---------------------------------------------------------------------------
# Configurazione ticker
# ---------------------------------------------------------------------------

# Aziende per l'analisi fondamentale (Modulo 3)
AZIENDE_TICKERS = ["TRN.MI", "RACE.MI", "MSFT", "GOOGL", "MC.PA"]

# Indici benchmark (FTSE MIB sostituito da S&P 500, MSCI World, ecc.)
BENCHMARK_TICKERS = ["^GSPC", "^STOXX50E", "URTH", "EEM"]

# Periodo dati: 5 anni storici
PERIODO = "5y"

# Cartella di destinazione: stessa directory di questo script
CARTELLA = Path(__file__).parent


# ---------------------------------------------------------------------------
# Funzione di pulizia nome file
# ---------------------------------------------------------------------------

def ticker_to_filename(ticker: str) -> str:
    """
    Converte un ticker in un nome file valido.
    Sostituisce i caratteri non ammessi nei nomi file:
      - Il punto '.' diventa '_'  (es. TRN.MI  -> TRN_MI)
      - Il cappello '^' viene rimosso  (es. ^GSPC -> GSPC)
    """
    clean = ticker.replace("^", "").replace(".", "_")
    return clean


# ---------------------------------------------------------------------------
# Funzione di download e salvataggio singolo ticker
# ---------------------------------------------------------------------------

def scarica_e_salva(ticker: str, prefisso: str = "") -> pd.DataFrame | None:
    """
    Scarica i dati storici di un singolo ticker da Yahoo Finance
    e li salva come CSV in CARTELLA.

    Parametri
    ----------
    ticker   : codice Yahoo Finance (es. "TRN.MI")
    prefisso : stringa da anteporre al nome file (es. "benchmark_")

    Ritorna
    -------
    DataFrame con i dati scaricati, oppure None in caso di errore.
    """
    nome_file = f"{prefisso}{ticker_to_filename(ticker)}.csv"
    percorso = CARTELLA / nome_file

    print(f"  Scaricando {ticker}...", end=" ", flush=True)

    try:
        df = yf.download(ticker, period=PERIODO, auto_adjust=True, progress=False)

        if df.empty:
            print(f"ATTENZIONE: nessun dato ricevuto per {ticker}")
            return None

        # Salva con indice (colonna Date)
        df.to_csv(percorso)

        date_min = df.index.min().strftime("%Y-%m-%d")
        date_max = df.index.max().strftime("%Y-%m-%d")
        print(f"OK  ({len(df)} righe, {date_min} -> {date_max})  -> {nome_file}")

        return df

    except Exception as e:
        print(f"ERRORE: {e}")
        return None


# ---------------------------------------------------------------------------
# Funzione principale
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("Turin Wealth Advisory - Generazione cache dati")
    print("=" * 60)
    print(f"Cartella destinazione: {CARTELLA}")
    print(f"Periodo: {PERIODO}\n")

    # Dizionario per raccogliere le serie di prezzi di chiusura
    close_prices: dict[str, pd.Series] = {}

    # ------------------------------------------------------------------
    # 1. Download dati aziende
    # ------------------------------------------------------------------
    print("--- AZIENDE ---")
    for ticker in AZIENDE_TICKERS:
        df = scarica_e_salva(ticker, prefisso="")
        if df is not None and "Close" in df.columns:
            # Estrae la serie Close e la rinomina con il ticker pulito
            close_prices[ticker_to_filename(ticker)] = df["Close"]

    # ------------------------------------------------------------------
    # 2. Download dati benchmark
    # ------------------------------------------------------------------
    print("\n--- BENCHMARK ---")
    for ticker in BENCHMARK_TICKERS:
        df = scarica_e_salva(ticker, prefisso="benchmark_")
        if df is not None and "Close" in df.columns:
            # Prefisso nel nome colonna per distinguerli dagli altri
            close_prices[f"bench_{ticker_to_filename(ticker)}"] = df["Close"]

    # ------------------------------------------------------------------
    # 3. File combinato: solo prezzi di chiusura, tutti i titoli
    # ------------------------------------------------------------------
    print("\n--- FILE COMBINATO ---")
    if close_prices:
        tutti = pd.DataFrame(close_prices)

        # Allinea gli indici (le date possono variare per mercati diversi)
        tutti.index.name = "Date"

        percorso_combinato = CARTELLA / "tutti_i_titoli.csv"
        tutti.to_csv(percorso_combinato)

        date_min = tutti.index.min().strftime("%Y-%m-%d")
        date_max = tutti.index.max().strftime("%Y-%m-%d")
        print(
            f"  Salvato tutti_i_titoli.csv  "
            f"({len(tutti)} righe, {len(tutti.columns)} colonne, "
            f"{date_min} -> {date_max})"
        )
        print(f"  Colonne: {list(tutti.columns)}")
    else:
        print("  ATTENZIONE: nessun dato disponibile per il file combinato.")

    # ------------------------------------------------------------------
    # 4. Riepilogo finale
    # ------------------------------------------------------------------
    print("\n" + "=" * 60)
    csv_files = sorted(CARTELLA.glob("*.csv"))
    print(f"Riepilogo: {len(csv_files)} file CSV presenti in {CARTELLA.name}/")
    for f in csv_files:
        size_kb = f.stat().st_size / 1024
        print(f"  {f.name:<35}  {size_kb:6.1f} KB")
    print("=" * 60)
    print("Cache completata. I notebook possono ora funzionare offline.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    main()
