# PID-Checker

A small web app that checks **P&ID drawings** against the project's engineering
lists and flags anything that appears on a drawing but is **missing from the
matching list**.

It reads every tag off the P&ID PDF and reconciles:

- **Equipment** (`AREA-TYPE-NNN`, e.g. `121-CV-014`) against the **MEL**
- **Valves** (sized `50V11A` and pattern `120-VV-003`) against the **Valve List**
- **Lines** (`SIZE-SERV-SPEC-NNN`, e.g. `25-PW-S31-027`) against the **Line List**

Anything tagged on a drawing whose type isn't tracked (e.g. CCTV cameras `CA`)
is also reported as a discrepancy.

## How to run the web app (the easy way)

You need this **one-time setup**:

1. **Install Python** from https://www.python.org/downloads/ (tick *"Add Python
   to PATH"* during install).
2. Open a terminal **in this project folder**:
   - Windows: type `cmd` in the folder's address bar and press Enter.
   - Mac: right-click the folder → *New Terminal at Folder*.
3. Install the libraries it needs (once):
   ```
   pip install -r requirements.txt
   ```

Then, **every time you want to use it**:

```
python app.py
```

You'll see a line like `Running on http://127.0.0.1:5000`. Open that address in
your web browser, **drag in the four files** (P&ID PDF, Line List, Valve List,
MEL), and click **Run consistency check**. You get the results on screen and a
**Download Excel report** button. Press `Ctrl+C` in the terminal to stop it.

## Prefer the command line?

The same analysis is also a script. Order of files: PDF, Line List, Valve List,
MEL, then the report name to create:

```
python check_project.py PID.pdf Line_List.xlsx Valve_List.xlsx MEL.xlsx report.xlsx
```

## What the report contains

- **Summary** — counts of equipment / valves / lines found and total discrepancies.
- **Discrepancies** — the items on a P&ID that are missing from a list (the point).
- **Full PID Inventory** — every tag found, so you can confirm nothing was missed.

## Adjusting it to other projects

The tag formats are calibrated to the 2233 Kathleen Valley drawings. If your
drawings use different conventions, the patterns live near the top of
`check_project.py` (and `config.py`) and can be adjusted there.

## Tests

```
python tests/test_check_project.py
```

Generates tiny synthetic files in the project's tag style and checks that the
reconciliation flags exactly the planted "on-drawing-but-not-in-list" items.

## Notes

- P&IDs are read as native PDF text (fast, exact). Scanned/raster drawings would
  need OCR (`tesseract`), which isn't required for native-text PDFs like these.
- Valves are matched by their physical tag. Instrument/actuated valve loop tags
  (`YV-`/`SV-` shown as instrument balloons) are out of scope for now.
```
