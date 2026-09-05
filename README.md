# Dynamic Report Automation — MVP

## Current goal
Load a DOCX template, replace controlled placeholders, and generate a new DOCX.

## Run
1. Install Python 3.10+
2. `pip install -r requirements.txt`
3. `python generate_report.py`

The generated file will appear in `output/generated_report.docx`.

## Next stages
- Preserve formatting at run level
- Dynamically expand long sections
- Dynamic tables
- Image insertion
- Page overflow detection
- Layout validation
- Streamlit UI
- Optional AI content structuring
