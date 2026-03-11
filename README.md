# Valuation Model → Output Transposer (Local)

Local Streamlit tool that:
1) Lets the user upload an **Input valuation model** (xlsm/xlsx)
2) Asks 16 filter questions defined directly in `src/filters.py`
3) Validates filters marked as required in code
4) Scans `Projections` columns **E:U** for columns where:
   - Row 6 ("Last Month of Period") == "December"
   - Row 8 ("Status") == "Actual"
5) For each qualifying column:
   - Transposes `Projections` rows **6:288** from that column into one output row, starting at **BD**
   - Fills static fields **A:BC** using Start Page cells + filter answers + (optional) lat/long geocode
   - Inserts a NEW row at **row 2** in the destination output workbook (Option A)

## Quick start
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app/streamlit_app.py
```

## Configuration
Edit `configs/settings.json`:
- `destination_workbook_path`: path to the central output workbook where rows are inserted at row 2
- `destination_sheet_name`: default "Output"
- `geocoding`: optional (provider/url/api_key)

To change filter labels, types, required flags, or dropdown options, edit `HARDCODED_FILTER_QUESTIONS` in `src/filters.py`.

## Notes / limitations
- This uses **openpyxl**. If your input `.xlsm` contains formulas that were not calculated & saved, values may be stale.
  Best practice: open the model in Excel, let it calculate, save, then upload.
- Geocoding is implemented as a **stub** by default; you can plug in a provider later.
