# OSA Chart Auto-Fill

Fills the **50-400mA** and **80-400mA** chart Excel templates from OSA raw-data TXT files
(e.g. `48226H40040_A2026100620533560886.txt`).

## Web app (Streamlit)
Upload one or more TXT files → download `<TXT>-Chart-50-400mA.xlsx` and `<TXT>-Chart-80-400mA.xlsx`
(several files are bundled into one zip). You can also upload a different template in the sidebar.

Run locally:
```
pip install -r requirements.txt
streamlit run app.py
```

## Command line (Windows / no Streamlit needed)
```
python fill_chart_templates.py RAW.txt [RAW2.txt ...]      # or a folder of TXT files
python fill_chart_templates.py                              # opens a file picker
python fill_chart_templates.py RAW.txt --t50 T50.xlsx --t80 T80.xlsx --out D:\Output
```

## What it does
- **RawData** sheet: replaced with the TXT data in the Power Query layout (`Temp_-5` → `Temp`, `-5`).
- **Mapping** sheet: formulas untouched; cached values recalculated; Excel fully recalculates on open.
- **Charts**: SN in the 4 chart titles updated; charts, styles and filters otherwise kept.

## Files
| File | Purpose |
|---|---|
| `app.py` | Streamlit web page |
| `fill_chart_templates.py` | core logic + command-line tool (standard library only) |
| `templates/Template-50-400mA.xlsx` | 50-400 mA template (36 currents × 4 CH × 17 temps) |
| `templates/Template-80-400mA.xlsx` | 80-400 mA template (33 currents × 4 CH × 17 temps) |
