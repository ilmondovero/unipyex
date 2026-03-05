# Python per Excel - Lezione Pratica (6h)

## Turin Wealth Advisory - "Dietro le quinte dei fogli Excel"

Materiale didattico per la lezione su Python applicato a Excel, parte del corso universitario di Excel per la facolta di Economia.

---

## Prerequisiti

- **Python 3.10+** installato
- **Excel** installato (necessario per xlwings)
- Conoscenza base di Python (if/for/funzioni)
- Aver completato almeno i Moduli A, B, C, D del corso Excel

## Setup rapido

```bash
# 1. Clona o scarica questa cartella
# 2. Apri un terminale nella cartella "lezione/"

# 3. Crea un ambiente virtuale (consigliato)
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # Mac/Linux

# 4. Installa le dipendenze
pip install -r requirements.txt

# 5. Avvia Jupyter
jupyter notebook
```

## Struttura della lezione

| Blocco | Durata | Notebook | Argomenti |
|--------|--------|----------|-----------|
| 1 | 2h | `01_dal_web_al_dataframe.ipynb` | yfinance, pandas, rendimenti |
| 2 | 2h | `02_dal_dataframe_a_excel.ipynb` | xlwings, formattazione, formule |
| 3 | 2h | `03_pipeline_completa.ipynb` | Automazione end-to-end |

## Materiale di supporto

- `slides/index.html` - Slide della lezione (aprire nel browser, navigare con le frecce)
- `cheatsheet/cheatsheet_python_excel.html` - Riferimento rapido (stampabile)
- `dati_cache/` - Dati offline di fallback (se la rete non e disponibile)

## Modalita offline

Se la connessione internet non e disponibile in aula:

```bash
# Prima della lezione, con connessione attiva:
cd dati_cache
python genera_cache.py
```

I notebook rileveranno automaticamente l'assenza di rete e caricheranno i dati dalla cache.

## Collegamento con il progetto Turin Wealth

I notebook importano direttamente da `../scripts/`:
- `tw_config.py` - Colori, costanti, dati del progetto
- `tw_utils.py` - Funzioni helper per xlwings

Questo permette agli studenti di vedere il codice reale usato per generare i file Excel del corso.

## Troubleshooting

| Problema | Soluzione |
|----------|-----------|
| `ModuleNotFoundError: xlwings` | `pip install xlwings` |
| xlwings non trova Excel | Verificare che Excel sia installato e chiuso |
| yfinance timeout | Usare la cache offline (vedi sopra) |
| Errore `Formula2Local` | Aggiornare xlwings: `pip install --upgrade xlwings` |
| Caratteri unicode (stelle) non visibili | Impostare encoding UTF-8 nel terminale |
