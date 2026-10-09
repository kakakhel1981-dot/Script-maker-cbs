# MSISDN Format Converter

A simple Streamlit app that reads an Excel file with MSISDN, IMSI, and date/time columns and creates one pipe-delimited output record per valid row.

## Files
- `app.py` — Streamlit application
- `requirements.txt` — dependencies
- `sample_input.xlsx` — sample workbook for testing

## Input columns
Use columns named `MSISDN`, `IMSI`, and `DATETIME` (or select the correct columns in the app). The date/time may be an Excel date/time or a value such as `2026-08-27 16:23:55`.

## Output format
`MSISDN|IMSI|MSISDN_abc|MSISDN_YYYYMMDDHHMMSS|MSISDN_YYYYMMDDHHMMSS||MSISDN_YYYYMMDDHHMMSS||||||||||`

The app uses Python formatting logic only; no LLM or API key is needed.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy with GitHub and Streamlit Community Cloud
1. Create a new GitHub repository.
2. Upload `app.py`, `requirements.txt`, and optionally `sample_input.xlsx` and `README.md`.
3. Open Streamlit Community Cloud and choose **Create app**.
4. Select your repository, branch, and `app.py` as the main file.
5. Deploy.

The app does not need secrets or API keys.
