# Corn Mill Performance Dashboard (Streamlit)

Interactive dashboard for corn mill sieve analysis data:
- PSD line graph (semi-log aperture mm vs cumulative % passing)
- Sieve distribution bar graph (% retained)
- Metrics table for all runs
- **Year-based PAES/PNS compliance** (Main Product Recovery + Losses only)
- **Compliance summary pie chart** (PASS vs FAIL)
- **Year filter toggle** (2018 group vs 2021 group)
- **Batch-level compliance score (0–100)**

## Compliance logic (as configured)
- **Run 01–08 (up to 2019)** → uses **2018** threshold: Main Product Recovery ≥64% AND Losses ≤5%
- **Run 09+ (2021–2025)** → uses **2021** threshold: Main Product Recovery ≥55% AND Losses ≤5%

## Run locally
1) Create environment + install
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

2) Run
```bash
streamlit run app.py
```

3) Open the app and **upload your Excel (.xlsx)** in the sidebar.
