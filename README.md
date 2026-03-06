# Valuation Model → Output Transposer (Local)

Local Streamlit tool that:
1) Lets the user upload an **Input valuation model** (xlsm/xlsx)
2) Asks 16 filter questions defined in the **Output template** (Filters tab)
3) Validates mandatory questions (row 8 contains 'x')
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
- `output_template_path`: path to your output template workbook (used to read Filters questions)
- `destination_workbook_path`: path to the central output workbook where rows are inserted at row 2
- `destination_sheet_name`: default "Output"
- `geocoding`: optional (provider/url/api_key)

## Notes / limitations
- This uses **openpyxl**. If your input `.xlsm` contains formulas that were not calculated & saved, values may be stale.
  Best practice: open the model in Excel, let it calculate, save, then upload.
- Geocoding is implemented as a **stub** by default; you can plug in a provider later.
