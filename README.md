# Corn Mill • Performance Dashboard (Streamlit)

A Streamlit dashboard inspired by the layout of the reference web dashboard.

## Features
- Upload-only: upload an Excel (.xlsx) and the dashboard renders immediately.
- Dashboard / Particle Distribution / Sieve Distribution / Library / Team tabs (layout similar to the reference).
- PSD curve: semi-log aperture (mm) vs cumulative % passing, smooth curve with **circle markers**.
- Sieve distribution: **one bar chart per run**, comparing all outlets in the same chart.
- Year-based compliance:
  - Run 01–08 → 2018 (Main Product Recovery ≥64%, Losses ≤5%)
  - Run 09+ → 2021 (Main Product Recovery ≥55%, Losses ≤5%)
- Failed runs section with reason: Recovery / Losses / Recovery + Losses
- Batch-level compliance score (0–100)

## Run locally
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

Then upload your Excel file in the sidebar.
