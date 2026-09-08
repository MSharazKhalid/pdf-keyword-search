# PDF Keyword Search

Searches a PDF for one phrase, or for every value in an Excel column,
and reports which page each hit is on. Uses file-picker windows.

## Setup

```
pip install -r requirements.txt
```

## Run

```
python search_pdf_pages.py
```

## What to edit before running

Nothing - it prompts you for the file and the search text.

## Never commit

Patient or client data (`.xlsx`, `.pdf`, `.csv`), `service_account.json`,
and chromedriver binaries. All are covered by `.gitignore`.
