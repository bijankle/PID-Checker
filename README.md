# PID-Checker

Checks **P&ID drawings** against the project's engineering lists and shows, side
by side, the value read from the **P&ID** next to the value from the **MEL /
Line List / Valve List** — with mismatches highlighted in red.

There are two ways to run it:

- **`pid-checker.html`** — the easy one. Double-click to open in your browser,
  drag in the P&ID PDF + the three Excel lists, get the report on screen and a
  downloadable Excel. Nothing to install (needs internet the first time it's
  opened, to load two helper libraries).
- **`check_project.py`** — the command-line engine (also the reference the
  browser tool is validated against).

## What it compares

Output is one block per item type, PID value beside list value:

| Block | Columns |
|---|---|
| **Equipment** | PID equip-no · MEL equip-no · PID description · MEL description |
| **Valves** | PID valve tag · list valve tag · PID size-code · list size-code |
| **Lines** | PID line-no · list line-no · PID spec · list spec · sheet(s) found on · list P&ID |

How the reading is made reliable:

- **Drawing number** is read from the **title block by position** (the `AREA-FP-NNN`
  word nearest the bottom of the sheet), so cross-references to other drawings
  aren't mistaken for the sheet's own number.
- **Equipment** counts as belonging to a sheet only if its tag appears **twice**
  (genuine equipment is labelled at the symbol *and* in the reference block;
  cross-references appear once). Tags labelled once anywhere are listed with a
  "verify" note rather than dropped.
- **Equipment descriptions** are read from the reference block; the **valve
  size-code and unique VV tag** are paired by proximity on the drawing.
- Equipment found on a sheet is also checked against the MEL's **`P&ID No.`**.
- **Valves** found on a sheet are checked against the valve list's **`P&ID`** — a
  valve drawn on a sheet the list doesn't assign it to is flagged.
- **Lines** are checked both ways: a line drawn on a sheet the line list doesn't
  list it on, *and* a line the list places on an uploaded sheet but which isn't
  actually drawn there ("listed but not drawn"). The pipe **spec** read from the
  line tag is compared to the line list's `SPEC` (a spec break along the line is
  not a mismatch; only a spec the list never mentions is flagged).

## Command-line use

```
pip install -r requirements.txt
python check_project.py PID.pdf Line_List.xlsx Valve_List.xlsx MEL.xlsx report.xlsx
```

## Tests

```
python tests/test_check_project.py
```

Generates tiny synthetic files in the project's tag style and checks the
side-by-side tables, the twice-rule, description extraction and valve pairing.

## Continuation-ribbon check

Off-page connectors (border flags) are read as `target-drawing + line-number`,
and each is verified against where that line actually appears: a ribbon that
names a drawing which doesn't carry the line is flagged. Shown on the
**Continuations** sheet.

## Sheets with no text

Sheets with no extractable text layer (scanned, or text saved as vector
outlines) can't be read reliably, so rather than guess they are listed on the
**Manual Review** sheet for a human to check. Best to re-export such a sheet
with real text when you can.

## Known limitations

- **Description extraction** is best-effort; a small number of equipment that
  sit in vendor-package blocks may read the wrong description.
- Instrument/actuated valve loop tags (`YV-`/`SV-` balloons) are out of scope.
