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

Run:  python check_project.py PID.pdf Line_List.xlsx Valve_List.xlsx MEL.xlsx [out.xlsx]
"""
from __future__ import annotations

import re
import sys

import fitz
import pandas as pd

# --- regexes for the on-drawing tag conventions ---------------------------
RE_EQ   = re.compile(r"\b(\d{3})-([A-Z]{2})-(\d{3})\b")
RE_VV   = re.compile(r"\b(\d{3})-VV-(\d{3})\b")
RE_VSZ  = re.compile(r"\b(\d{2,4})V(\d{2})([A-Z])\b")
RE_LINE = re.compile(r"\b(\d{2,4})-([A-Z]{2,3})-([A-Z]\d{2})-(\d{3})\b")
RE_DWG  = re.compile(r"\b(\d{3}-FP-\d{3})\b")

# Codes that are NOT mechanical equipment even though they fit AREA-XX-NNN.
NON_EQUIPMENT_CODES = {"FP"}  # drawing references; VV handled separately


def load_references(line_path, valve_path, mel_path):
    """Build the identifier sets each P&ID tag is checked against."""
    # MEL equipment + the set of valid equipment type codes.
    mel = pd.read_excel(mel_path, sheet_name="Equipment List", header=0, dtype=str)
    equipment, codes = set(), set()
    for v in mel["Equipment No."].dropna():
        m = RE_EQ.fullmatch(str(v).strip().upper())
        if m:
            equipment.add(m.group(0))
            codes.add(m.group(2))

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
            "serv_num": serv_num, "serv_spec_num": serv_spec_num}


def extract_page(text, codes):
    """Return dict of category -> set(tags) found in one page's text."""
    T = text.upper()
    eq, valves, lines, other = set(), set(), set(), set()
    for m in RE_LINE.finditer(T):
        lines.add(m.group(0))
    for m in RE_EQ.finditer(T):
        code, tag = m.group(2), m.group(0)
        if code in NON_EQUIPMENT_CODES:
            continue
        if code == "VV":
            valves.add(tag)
        elif code in codes:
            eq.add(tag)
        else:
            other.add(tag)   # tagged item whose type isn't a MEL code (CA, SP...)
    for m in RE_VSZ.finditer(T):
        valves.add(m.group(0))
    return {"Equipment": eq, "Valve": valves, "Line": lines, "Other": other}


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
    rows = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text = page.get_text("text")
            m = RE_DWG.search(text)
            dwg = m.group(1) if m else f"(page {page.number + 1})"
            found = extract_page(text, ref["codes"])
            for tag in sorted(found["Equipment"]):
                ok = tag in ref["equipment"]
                rows.append([dwg, "Equipment", tag, "MEL",
                             "IN LIST" if ok else "NOT IN MEL"])
            for tag in sorted(found["Valve"]):
                ok = tag in ref["valves"]
                rows.append([dwg, "Valve", tag, "Valve List",
                             "IN LIST" if ok else "NOT IN VALVE LIST"])
            for tag in sorted(found["Line"]):
                ok = line_in_list(tag, ref)
                rows.append([dwg, "Line", tag, "Line List",
                             "IN LIST" if ok else "NOT IN LINE LIST"])
            for tag in sorted(found["Other"]):
                # per scope decision: untracked tagged items are discrepancies
                rows.append([dwg, "Other", tag, "MEL", "NOT IN MEL (untracked type)"])

    df = pd.DataFrame(rows, columns=["P&ID", "Category", "Tag",
                                     "Checked Against", "Status"])
    disc = df[df["Status"] != "IN LIST"].reset_index(drop=True)
    summary = pd.DataFrame({
        "Metric": ["P&ID pages", "Equipment found", "Valves found", "Lines found",
                   "Other tagged items", "DISCREPANCIES (on P&ID, not in list)"],
        "Value": [df["P&ID"].nunique(),
                  int((df.Category == "Equipment").sum()),
                  int((df.Category == "Valve").sum()),
                  int((df.Category == "Line").sum()),
                  int((df.Category == "Other").sum()),
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
