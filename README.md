# PDF Keyword Search

Finds a phrase - or every value in an Excel column - inside a PDF and reports the
exact page of each hit.

![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white) ![PyMuPDF](https://img.shields.io/badge/PyMuPDF-1A73E8) ![License](https://img.shields.io/badge/License-MIT-2ea44f)

## What it does

Two modes:

- **Single phrase** - type it in, get back every page it appears on.
- **Bulk** - point it at an Excel column and it searches for all of those values in
  one pass, so you are not re-opening the same 400-page document once per term.

Opens file-picker windows for both the PDF and the spreadsheet, so there is nothing
to edit in the code before running it.

## Requirements

- Python 3.9+
- `PyMuPDF`, `openpyxl`

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python search_pdf_pages.py
```

No configuration needed - it prompts for the file and the search text.

## Notes on data

This repository contains **no client or patient data**. `.gitignore` already excludes
`.xlsx`, `.pdf` and `.csv` files, `service_account.json`, and chromedriver binaries -
keep it that way if you fork this.

## License

MIT © Muhammad Sharaz Khalid - see [LICENSE](LICENSE).
