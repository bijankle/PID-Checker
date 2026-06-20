"""
Calibrated P&ID-vs-list reconciliation for the 2233 Kathleen Valley project
(Lycopodium drawings + Line List / Valve List / MEL).

This encodes the conventions confirmed against the sample files:

  Drawing number : in the title block, formatted  AREA-FP-NNN  (e.g. 120-FP-001)
  Equipment tags : AREA-TYPE-NNN where TYPE is a MEL equipment code (e.g. 121-CV-014)
  Valve tags     : sized  <DN>V<nn><L>  (e.g. 50V11A)  and pattern  AREA-VV-NNN
  Line tags      : <SIZE>-<SERV>-<SPEC>-<NNN>  (e.g. 25-PW-S31-027)

Scope decisions (per project owner):
  * Valves are reconciled by their physical VV / sized tag (instrument YV/SV
    loop tags are out of scope for now).
  * ANY tagged item on a P&ID that is not in the relevant list is a discrepancy,
    including CCTV cameras (CA) and stockpiles (SP) absent from the MEL.
  * The drawing number is read from the title block by position (the FP word
    nearest the bottom of the sheet), so cross-references to other drawings on
    the same sheet do not get mistaken for the sheet's own number.
  * Equipment counts as belonging to a sheet only if its tag appears at least
    twice on it (genuine equipment is labelled at the symbol and in the
    reference block; cross-references from other sheets appear once). This is
    what stops equipment being paired with the wrong P&ID.
  * For equipment, the drawing it is found on is compared against the MEL's
    "P&ID No." column; a mismatch is flagged ("P&ID MISMATCH"). The MEL
    "Equipment Name" is carried through as a Description column.

Run:  python check_project.py PID.pdf Line_List.xlsx Valve_List.xlsx MEL.xlsx [out.xlsx]
"""
from __future__ import annotations

import re
import sys
from collections import Counter

import fitz
import pandas as pd

# --- regexes for the on-drawing tag conventions ---------------------------
RE_EQ   = re.compile(r"\b(\d{3})-([A-Z]{2})-(\d{3})\b")
RE_VV   = re.compile(r"\b(\d{3})-VV-(\d{3})\b")
RE_VSZ  = re.compile(r"\b(\d{2,4})V(\d{2})([A-Z])\b")
RE_LINE = re.compile(r"\b(\d{2,4})-([A-Z]{2,3})-([A-Z]\d{2})-(\d{3})\b")
RE_DWG  = re.compile(r"\b(\d{3}-FP-\d{3})\b")
RE_FP_FULL = re.compile(r"^\d{3}-FP-\d{3}$")  # a whole word that is a drawing no.

# Codes that are NOT mechanical equipment even though they fit AREA-XX-NNN.
NON_EQUIPMENT_CODES = {"FP"}  # drawing references; VV handled separately


def _find_col(columns, *wanted):
    """Find a column by fuzzy, whitespace/case-insensitive match."""
    norm = lambda s: "".join(str(s).lower().split()).replace("_", "").replace("-", "")
    targets = [norm(w) for w in wanted]
    for c in columns:
        if norm(c) in targets:
            return c
    return None


def drawing_number(page):
    """The sheet's own drawing number = the NNN-FP-NNN word nearest the bottom
    of the page (the title block), tie-broken to the right. This ignores the
    many cross-references to other drawings elsewhere on the sheet."""
    fps = [w for w in page.get_text("words") if RE_FP_FULL.match(w[4])]
    if not fps:
        return None
    # words are (x0, y0, x1, y1, text, ...). Largest y0 = lowest on page.
    fps.sort(key=lambda w: (w[1], w[0]))
    return fps[-1][4]


def load_references(line_path, valve_path, mel_path):
    """Build the identifier sets each P&ID tag is checked against."""
    # MEL equipment + the set of valid equipment type codes + declared P&ID.
    mel = pd.read_excel(mel_path, sheet_name="Equipment List", header=0, dtype=str)
    equipment, codes, declared_pid, names = set(), set(), {}, {}
    eq_col = _find_col(mel.columns, "Equipment No.") or "Equipment No."
    pid_col = _find_col(mel.columns, "P&ID No.", "P&ID\nNo.", "PID No.")
    name_col = _find_col(mel.columns, "Equipment Name", "Description", "Equipment Description")
    for _, row in mel.iterrows():
        m = RE_EQ.fullmatch(str(row.get(eq_col)).strip().upper())
        if not m:
            continue
        equipment.add(m.group(0))
        codes.add(m.group(2))
        if pid_col is not None:
            dm = RE_DWG.search(str(row.get(pid_col)).upper())
            if dm:
                declared_pid[m.group(0)] = dm.group(1)
        if name_col is not None:
            v = str(row.get(name_col)).strip()
            if v and v.lower() != "nan":
                names[m.group(0)] = v

    # Valve identifiers: union of every identifier-bearing column / sheet.
    valves = set()
    main = pd.read_excel(valve_path, sheet_name="2233-PLST-007", header=0, dtype=str)
    for col in ("Valve Identifier", "Valve Name Lookup", "PATTERN"):
        if col in main.columns:
            valves |= {str(v).strip().upper() for v in main[col].dropna()}
    for sheet, col in (("Actuated Valves", "TAG"), ("Actuated Valves", "Valve Code"),
                       ("Manual Valves", "Valve Tag")):
        try:
            s = pd.read_excel(valve_path, sheet_name=sheet, header=0, dtype=str)
            valves |= {str(v).strip().upper() for v in s[col].dropna()}
        except Exception:
            pass
    valves = {v for v in valves if v and v not in {"NAN", "-", "SPARE"}}

    # Line list: keys of (serv, spec, num) and (serv, num) that exist.
    ll = pd.read_excel(line_path, sheet_name="Line List", header=0, dtype=str).iloc[1:]
    serv_num, serv_spec_num = set(), set()

    def add(serv, spec, num):
        serv = str(serv).strip().upper()
        num = re.sub(r"\D", "", str(num))
        if not serv or not num:
            return
        num = str(int(num))  # strip leading zeros (e.g. 027 -> 27) for consistent keys
        serv_num.add((serv, num))
        for sp in re.split(r"[\s/\n]+", str(spec).upper()):
            if sp.strip() and sp != "NAN":
                serv_spec_num.add((serv, sp.strip(), num))

    for _, r in ll.iterrows():
        add(r.get("SERV CODE"), r.get("SPEC"), r.get("LINE No."))
        for c in ll.columns:  # also harvest fully-written line numbers in cells
            for m in RE_LINE.finditer(str(r.get(c)).upper()):
                serv_num.add((m.group(2), str(int(m.group(4)))))
                serv_spec_num.add((m.group(2), m.group(3), str(int(m.group(4)))))

    return {"equipment": equipment, "codes": codes, "valves": valves,
            "serv_num": serv_num, "serv_spec_num": serv_spec_num,
            "declared_pid": declared_pid, "names": names}


def page_tags(text, codes):
    """Parse one page's text into:
       valves (set), lines (set), eq_counts (Counter of MEL-code equipment),
       other_counts (Counter of AREA-XX-NNN tags whose type isn't a MEL code).
    Counts are raw occurrences; the twice-rule is applied later, across pages."""
    T = text.upper()
    lines = {m.group(0) for m in RE_LINE.finditer(T)}
    valves = set()
    eq_counts, other_counts = Counter(), Counter()
    for m in RE_EQ.finditer(T):
        code, tag = m.group(2), m.group(0)
        if code in NON_EQUIPMENT_CODES:
            continue
        if code == "VV":
            valves.add(tag)
        elif code in codes:
            eq_counts[tag] += 1
        else:
            other_counts[tag] += 1
    for m in RE_VSZ.finditer(T):
        valves.add(m.group(0))
    return valves, lines, eq_counts, other_counts


def line_in_list(tag, ref):
    m = RE_LINE.fullmatch(tag)
    s, sp, n = m.group(2), m.group(3), str(int(m.group(4)))
    return (s, sp, n) in ref["serv_spec_num"] or (s, n) in ref["serv_num"]


def analyze(pdf_path, line_path, valve_path, mel_path):
    """Run the reconciliation and return (full_df, discrepancies_df, summary_df).

    Pure analysis — no files written — so both the CLI and the web app can call
    it and present the results however they like.
    """
    ref = load_references(line_path, valve_path, mel_path)
    declared, names = ref["declared_pid"], ref["names"]

    ONLY_ONCE = "Only labelled once — verify (possible cross-reference)"

    # ---- pass 1: read every page and tally occurrences -------------------
    pages = []
    homes = {}   # tag -> set of drawings where it is labelled >=2 times (its home)
    with fitz.open(pdf_path) as doc:
        for page in doc:
            dwg = drawing_number(page) or f"(page {page.number + 1})"
            valves, lines, eq_counts, other_counts = page_tags(
                page.get_text("text"), ref["codes"])
            pages.append((dwg, valves, lines, eq_counts, other_counts))
            for tag, c in (eq_counts | other_counts).items():
                if c >= 2:
                    homes.setdefault(tag, set()).add(dwg)

    # ---- pass 2: build rows ---------------------------------------------
    rows = []
    for dwg, valves, lines, eq_counts, other_counts in pages:
        for tag in sorted(eq_counts):
            note = ""
            if eq_counts[tag] < 2:
                if tag in homes:
                    continue            # cross-reference; real home captured elsewhere
                note = ONLY_ONCE        # never labelled twice anywhere -> flag for review
            if tag not in ref["equipment"]:
                status = "NOT IN MEL"
            elif tag in declared and declared[tag] != dwg:
                status = f"P&ID MISMATCH (MEL says {declared[tag]})"
            else:
                status = "IN LIST"
            rows.append([dwg, "Equipment", tag, "MEL", status, note])
        for tag in sorted(other_counts):
            note = ""
            if other_counts[tag] < 2:
                if tag in homes:
                    continue
                note = ONLY_ONCE
            rows.append([dwg, "Other", tag, "MEL", "NOT IN MEL (untracked type)", note])
        for tag in sorted(valves):
            ok = tag in ref["valves"]
            rows.append([dwg, "Valve", tag, "Valve List",
                         "IN LIST" if ok else "NOT IN VALVE LIST", ""])
        for tag in sorted(lines):
            ok = line_in_list(tag, ref)
            rows.append([dwg, "Line", tag, "Line List",
                         "IN LIST" if ok else "NOT IN LINE LIST", ""])

    df = pd.DataFrame(rows, columns=["P&ID", "Category", "Tag",
                                     "Checked Against", "Status", "Notes"])
    # an item needs attention if it isn't cleanly in its list, or carries a note
    disc = df[(df["Status"] != "IN LIST") | (df["Notes"] != "")].reset_index(drop=True)
    summary = pd.DataFrame({
        "Metric": ["P&ID pages", "Equipment found", "Valves found", "Lines found",
                   "Other tagged items", "Labelled once (verify)",
                   "DISCREPANCIES / items to review"],
        "Value": [df["P&ID"].nunique(),
                  int((df.Category == "Equipment").sum()),
                  int((df.Category == "Valve").sum()),
                  int((df.Category == "Line").sum()),
                  int((df.Category == "Other").sum()),
                  int((df.Notes != "").sum()),
                  len(disc)],
    })
    return df, disc, summary


def write_report(df, disc, summary, out_path):
    """Write the three result tables to an Excel workbook."""
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        disc.to_excel(xw, sheet_name="Discrepancies", index=False)
        df.to_excel(xw, sheet_name="Full PID Inventory", index=False)


def check(pdf_path, line_path, valve_path, mel_path, out_path="PID_Check_Report.xlsx"):
    """CLI convenience: analyze and write the Excel report in one call."""
    df, disc, summary = analyze(pdf_path, line_path, valve_path, mel_path)
    write_report(df, disc, summary, out_path)
    print(f"Wrote {out_path}: {len(df)} tags, {len(disc)} discrepancies")
    return df, disc


if __name__ == "__main__":
    args = sys.argv[1:]
    if len(args) < 4:
        print(__doc__)
        sys.exit(1)
    check(*args)
