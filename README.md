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
- `output_storage`: set `"type": "github"` and configure `github_repo` plus `github_excel_path` to store `OUTPUT.xlsx` in GitHub instead of a local path

For GitHub-backed output storage, set `GITHUB_TOKEN` in your environment before running the app.

To change filter labels, types, required flags, or example dropdown options, edit `HARDCODED_FILTER_QUESTIONS` in `src/filters.py`.
User-added dropdown values are saved in `configs/filter_options.json` and will appear in later sessions.

## Notes / limitations
- This uses **openpyxl**. If your input `.xlsm` contains formulas that were not calculated & saved, values may be stale.
  Best practice: open the model in Excel, let it calculate, save, then upload.
- Preview and upload now geocode the extracted address using Nominatim to populate `lat` and `long`. If geocoding fails, the app will continue with empty coordinates and show a warning.

## Output mapping template
- Run `python tools/generate_output_mapping_template.py` to regenerate [docs/output_mapping_template.xlsx](/workspaces/Val-Model-Transposer/docs/output_mapping_template.xlsx), which documents each output column, its title, its source, and the expected inserted-row result.
