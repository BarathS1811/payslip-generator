# Payslip Generator

Upload a payroll Excel sheet, get back a print-ready PDF with 4 payslips per A4 page.
No calculations are performed — every value on the payslip is taken directly from
your Excel columns.

See the full step-by-step deployment guide in the chat where this was built, or
follow these steps:

## Deploy in 3 steps (no coding required)

1. **Create a GitHub repository** and upload these files (see main instructions).
2. **Create a Render account** at https://render.com and connect your GitHub repo.
3. **Deploy** — Render will detect the `Procfile` and `requirements.txt` automatically.

## Running locally (optional, for testing)

```bash
pip install -r requirements.txt
python app.py
```

Then open http://localhost:5000 in your browser.

## Files

- `app.py` — the whole application (upload handling, Excel reading, PDF generation)
- `templates/index.html` — the upload page
- `requirements.txt` — Python packages needed
- `Procfile` — tells Render how to start the app
