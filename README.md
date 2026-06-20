# PID-Checker

A web app that reconciles **P&ID drawings** against engineering **lists**
(line list, valve list, tie-in list, mechanical equipment list, …) and flags
inconsistencies between them.

For every page of the P&ID PDF it:

1. reads the **drawing / P&ID number** from the title block (falling back to a
   full-page scan), and
2. finds every **line / valve / tie-in / equipment** tag on the page,

then attributes each tag to its drawing and cross-checks against the lists.

## What it flags

| Category | Meaning | Severity |
|---|---|---|
| `in_list_not_on_pid` | Item is in a list but on no drawing | error |
| `on_pid_not_in_list` | Tag is on a drawing but missing from its list | error |
| `wrong_declared_drawing` | List names a P&ID, but item is on a different one | error |
| `duplicate_in_list` | Same identifier appears twice in one list | warning |
| `missing_drawing_number` | A page's title block yielded no number | warning |
| `on_multiple_drawings` | Tag appears on several drawings (often fine) | info |

Results are shown in the browser and downloadable as an Excel report
(Summary + Findings sheets).

## Running it

```bash
pip install -r requirements.txt
python app.py
# open http://127.0.0.1:5000
```

Upload a P&ID PDF (required) plus any of the lists you have.

### OCR (optional)

P&IDs that are native/vector PDFs are read directly — no OCR needed. Scanned
pages fall back to OCR, which requires the **tesseract** binary in addition to
the Python packages:

```bash
# Debian/Ubuntu
sudo apt-get install tesseract-ocr
# macOS
brew install tesseract
```

If tesseract is absent, scanned pages are simply reported as un-OCR'd rather
than failing.

## Calibrating to your documents

Everything format-specific lives in **`config.py`**:

- `TITLE_BLOCK` — region of the page searched for the drawing number.
- `DRAWING_NUMBER_PATTERN` / `DRAWING_NUMBER_LABELS` — how the number looks.
- `TAG_PATTERNS` — regexes for line/valve/tie-in/equipment tags as drawn.
- `LIST_SCHEMAS` — which spreadsheet columns hold each identifier + attributes.

These ship with broad defaults. Once real sample files are available they get
tuned to the project's actual conventions.

## Tests

```bash
python tests/test_pipeline.py     # or: pytest
```

The test generates a synthetic P&ID + line list and asserts the planted
inconsistencies are detected.

## Project layout

```
app.py               Flask web app (upload + report)
config.py            All format-specific configuration (calibrate here)
pidchecker/
  pdf_extractor.py   Per-page drawing number + tag extraction (PyMuPDF, OCR fallback)
  excel_loader.py    Robust .xls/.xlsx loading with header + column detection
  reconciler.py      Cross-check engine producing findings
  report.py          Excel report generation
templates/, static/  Web UI
tests/               End-to-end test on synthetic data
```
