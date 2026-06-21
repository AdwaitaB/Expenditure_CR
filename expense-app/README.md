# AB Group — Expense Analyser

Internal finance tool for AB Group. Phase 1.

## Features
- Upload UPI bank statement Excel files (single or multi-sheet)
- Auto-detect expense heads and map to main categories
- Prompt user to manually categorise missing expense heads
- Add cash expenses manually with dropdown category selection
- Spending analysis with pie chart, bar chart, monthly trend
- Filter by month and search transactions

## Run Locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy Free on Streamlit Community Cloud
1. Push this folder to a GitHub repo
2. Go to https://share.streamlit.io
3. Connect your GitHub, select the repo, set `app.py` as the entry point
4. Click Deploy — done!

## File Structure
```
expense-app/
├── app.py            ← Main app
├── requirements.txt  ← Dependencies
└── README.md
```
