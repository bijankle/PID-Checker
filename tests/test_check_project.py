"""
End-to-end test of the calibrated checker on small synthetic files that mimic
the real project's conventions, so the logic is verifiable without shipping the
actual (confidential) drawings.

Covers the per-type side-by-side tables (PID value beside list value), the
twice-rule for equipment, description extraction from the reference block, and
valve dual-code pairing.
"""
import os
import sys

import fitz
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from check_project import analyze  # noqa: E402


def _make_pdf(path):
    doc = fitz.open()

    # --- page 1: drawing 120-FP-001 ---
    p1 = doc.new_page(width=842, height=595)
    p1.insert_text((50, 40), "130-FP-099")     # cross-reference to another drawing
    p1.insert_text((50, 60), "121-DC-050")     # ONCE here; real home is page 2
    # genuine equipment: labelled twice (reference block + symbol). The reference
    # block occurrence (left column) has the description on the line below it.
    p1.insert_text((50, 90), "121-CV-014")
    p1.insert_text((50, 108), "PRIMARY CRUSHER")
    p1.insert_text((50, 160), "121-PP-002")
    p1.insert_text((50, 178), "CRUSHING WATER PUMP")
    p1.insert_text((420, 300), "121-CV-014")   # symbol occurrences (no description)
    p1.insert_text((420, 360), "121-PP-002")
    # a valve: unique VV tag with its sized code right beside it
    p1.insert_text((300, 250), "120-VV-003")
    p1.insert_text((340, 250), "25V41A")
    # a line on the sheet
    p1.insert_text((250, 200), "25-PW-S31-027")
    p1.insert_text((700, 575), "120-FP-001")   # title-block drawing number

    # --- page 2: 130-FP-002, the real home of 121-DC-050 (twice) ---
    p2 = doc.new_page(width=842, height=595)
    for y in (90, 320):
        p2.insert_text((50, y), "121-DC-050")
    p2.insert_text((50, 108), "DUST COLLECTOR")
    p2.insert_text((700, 575), "130-FP-002")

    doc.save(path)
    doc.close()


def _make_lists(tmp_path):
    mel = tmp_path / "mel.xlsx"
    pd.DataFrame({
        "Equipment No.": ["121-CV-014", "121-PP-002", "121-DC-050"],
        "P&ID\nNo.": ["120-FP-001", "120-FP-001", "130-FP-002"],
        "Equipment Name": ["Primary Crusher", "Slurry Pump", "Dust Collector"],  # PP-002 differs
    }).to_excel(mel, sheet_name="Equipment List", index=False)

    valve = tmp_path / "valve.xlsx"
    with pd.ExcelWriter(valve, engine="openpyxl") as xw:
        pd.DataFrame({"Valve Identifier": ["V41A"], "Valve Name Lookup": ["25V41A"],
                      "PATTERN": ["120-VV-003"]}).to_excel(
            xw, sheet_name="2233-PLST-007", index=False)
        pd.DataFrame({"TAG": [], "Valve Code": []}).to_excel(
            xw, sheet_name="Actuated Valves", index=False)
        pd.DataFrame({"Valve Tag": []}).to_excel(
            xw, sheet_name="Manual Valves", index=False)

    line = tmp_path / "line.xlsx"
    pd.DataFrame([
        ["sub", "", "", "", ""],
        ["", "PW", "S31", "027", "120-FP-001"],   # zero-padded; P&ID matches the sheet
    ], columns=["Identifier", "SERV CODE", "SPEC", "LINE No.", "P&ID"]).to_excel(
        line, sheet_name="Line List", index=False)
    return line, valve, mel


def _row(df, col, val):
    r = df[df[col] == val]
    return r.iloc[0] if len(r) else None


def test_analyze(tmp_path):
    pdf = tmp_path / "pid.pdf"
    _make_pdf(str(pdf))
    line, valve, mel = _make_lists(tmp_path)

    res = analyze(str(pdf), str(line), str(valve), str(mel))
    eq = res["tables"]["Equipment"].fillna("")
    eqm = res["masks"]["Equipment"]
    vl = res["tables"]["Valves"].fillna("")
    ln = res["tables"]["Lines"].fillna("")

    # --- Equipment: side-by-side + twice-rule + description ---
    assert list(eq.columns) == ["P&ID", "PID Equip No", "MEL Equip No",
                                "PID Description", "MEL Description", "Notes"]
    cv = _row(eq, "PID Equip No", "121-CV-014")
    assert cv["P&ID"] == "120-FP-001" and cv["MEL Equip No"] == "121-CV-014"
    assert cv["PID Description"] == "PRIMARY CRUSHER" and cv["MEL Description"] == "Primary Crusher"

    # PP-002: description disagrees (PID "CRUSHING WATER PUMP" vs MEL "Slurry Pump") -> red
    pp_idx = eq.index[eq["PID Equip No"] == "121-PP-002"][0]
    assert bool(eqm.loc[pp_idx, "MEL Description"]) is True

    # 121-DC-050 placed on its real home (page 2), not on page 1 where it's a one-off
    dc = _row(eq, "PID Equip No", "121-DC-050")
    assert dc["P&ID"] == "130-FP-002"
    assert not len(eq[(eq["PID Equip No"] == "121-DC-050") & (eq["P&ID"] == "120-FP-001")])

    # --- Valves: dual codes paired and matched ---
    v = _row(vl, "PID Valve Tag", "120-VV-003")
    assert v["List Valve Tag"] == "120-VV-003"
    assert v["PID Size Code"] == "25V41A" and v["List Size Code"] == "25V41A"

    # --- Lines: side-by-side + list P&ID ---
    li = _row(ln, "PID Line No", "25-PW-S31-027")
    assert li["List Line No"] == "25-PW-S31-027"
    assert "120-FP-001" in li["List P&ID"]


if __name__ == "__main__":
    import tempfile, pathlib
    d = pathlib.Path(tempfile.mkdtemp())
    test_analyze(d)
    print("OK")
