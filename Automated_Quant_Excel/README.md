# Financial PDF → Excel Pipeline
## ABBYY Cloud OCR SDK + Gemini API

## Setup (one time)
```bash
pip install -r requirements.txt
```

## How to Connect ABBYY Cloud OCR SDK
1. Go to https://cloud.ocrsdk.com → Sign in
2. Click "New Application" → Give it a name
3. You receive Application ID + Application Password via email
4. Paste both into main.py under CONFIGURATION section

## How to get Gemini API Key
1. Go to https://aistudio.google.com/app/apikey
2. Click "Create API Key"
3. Paste into main.py under CONFIGURATION section

## Running (Fresh PDF)
1. Set PDF_PATH to your file in ~/Downloads/
2. Make sure EXISTING_EXCEL = None
3. Run: python main.py

## Running (Year-on-Year Update)
1. Set PDF_PATH to the NEW year PDF
2. Set EXISTING_EXCEL to the path of your OLD year Excel file
3. Uncomment gemini_extract_update() function in main.py
4. Change EXISTING_EXCEL = None to EXISTING_EXCEL = "path/to/old.xlsx"
5. Run: python main.py

## Output
- xyz_financial_output.xlsx  → formatted Excel with all tables
- xyz_financial_output_gemini_raw.json → raw Gemini response (for debugging)
