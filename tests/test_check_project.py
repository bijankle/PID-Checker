"""
End-to-end test of the calibrated checker on small synthetic files that mimic
the real project's conventions, so the logic is verifiable without shipping the
actual (confidential) drawings.

Covers: tags found on the drawing vs the lists; position-based drawing-number
detection (title block at the bottom, ignoring cross-references); the
"drawn here but the MEL says a different P&ID" cross-check; the twice-rule
(genuine equipment is labelled twice, cross-references once); and the
"only labelled once anywhere -> flag with a note" behaviour.
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
    p1.insert_text((50, 40), "130-FP-099")    # cross-reference to another drawing
    p1.insert_text((50, 60), "121-DC-050")    # ONCE here, but its real home is page 2
    p1.insert_text((50, 75), "121-CV-077")    # ONCE and no home anywhere -> flagged with a note
    for y in (90, 320):                       # genuine items: labelled twice
        p1.insert_text((50, y), "121-CV-014")     # equipment, MEL home = this sheet
        p1.insert_text((150, y), "121-PP-002")    # equipment, MEL home = a DIFFERENT sheet
        p1.insert_text((250, y), "121-CV-099")    # equipment NOT in MEL
        p1.insert_text((350, y), "120-CA-202")    # camera: untracked type
    p1.insert_text((50, 150), "120-VV-003")   # valve in list (single label is normal)
    p1.insert_text((50, 170), "50V11A")       # sized valve in list
    p1.insert_text((50, 190), "25-PW-S31-027")  # line in list
    p1.insert_text((50, 210), "25-PW-S31-999")  # line NOT in list
    p1.insert_text((700, 575), "120-FP-001")  # title-block drawing number (bottom-right)

    # --- page 2: drawing 130-FP-002, the real home of 121-DC-050 (twice) ---
    p2 = doc.new_page(width=842, height=595)
    for y in (90, 320):
        p2.insert_text((50, y), "121-DC-050")
    p2.insert_text((700, 575), "130-FP-002")

    doc.save(path)
    doc.close()


def _make_lists(tmp_path):
    mel = tmp_path / "mel.xlsx"
    pd.DataFrame({
        "Equipment No.": ["121-CV-014", "121-PP-002", "121-DC-050"],
        "P&ID\nNo.": ["120-FP-001", "130-FP-002", "130-FP-002"],
        "Equipment Name": ["Primary Crusher", "Crushing Water Pump", "Dust Collector"],
    }).to_excel(mel, sheet_name="Equipment List", index=False)

    valve = tmp_path / "valve.xlsx"
    with pd.ExcelWriter(valve, engine="openpyxl") as xw:
        pd.DataFrame({"Valve Identifier": ["V11A"], "Valve Name Lookup": ["50V11A"],
                      "PATTERN": ["120-VV-003"]}).to_excel(
            xw, sheet_name="2233-PLST-007", index=False)
        pd.DataFrame({"TAG": [], "Valve Code": []}).to_excel(
            xw, sheet_name="Actuated Valves", index=False)
        pd.DataFrame({"Valve Tag": []}).to_excel(
            xw, sheet_name="Manual Valves", index=False)

    line = tmp_path / "line.xlsx"
    df = pd.DataFrame([
        ["sub", "", "", ""],
        ["", "PW", "S31", "027"],   # zero-padded on purpose: must still match 25-PW-S31-027
    ], columns=["Identifier", "SERV CODE", "SPEC", "LINE No."])
    df.to_excel(line, sheet_name="Line List", index=False)
    return line, valve, mel


def _row(df, tag):
    r = df[df["Tag"] == tag]
    return r.iloc[0] if len(r) else None


def test_analyze(tmp_path):
    pdf = tmp_path / "pid.pdf"
    _make_pdf(str(pdf))
    line, valve, mel = _make_lists(tmp_path)

    df, disc, summary = analyze(str(pdf), str(line), str(valve), str(mel))
    df = df.fillna("")

    status = dict(zip(df["Tag"], df["Status"]))
    notes = dict(zip(df["Tag"], df["Notes"]))

    # twice-rule: a single occurrence with a real home elsewhere is dropped here,
    # and the item is correctly placed on its home sheet (page 2).
    dc = _row(df, "121-DC-050")
    assert dc is not None and dc["P&ID"] == "130-FP-002" and dc["Notes"] == ""
    assert not len(df[(df["Tag"] == "121-DC-050") & (df["P&ID"] == "120-FP-001")])

    # a single occurrence with no home anywhere is listed, with a note
    assert "121-CV-077" in status
    assert "once" in notes["121-CV-077"].lower()

    # plain discrepancies
    assert status["121-CV-099"] == "NOT IN MEL"
    assert status["25-PW-S31-999"] == "NOT IN LINE LIST"
    assert status["120-CA-202"].startswith("NOT IN MEL")
    assert "P&ID MISMATCH" in status["121-PP-002"]   # MEL home is 130-FP-002

    # correct items are not flagged
    assert status["121-CV-014"] == "IN LIST"
    assert status["120-VV-003"] == "IN LIST"
    assert status["50V11A"] == "IN LIST"
    assert status["25-PW-S31-027"] == "IN LIST"      # zero-padded list entry still matches

    # position-based numbering picked the bottom title block, not the top cross-ref
    assert "130-FP-099" not in set(df["P&ID"])


if __name__ == "__main__":
    import tempfile, pathlib
    d = pathlib.Path(tempfile.mkdtemp())
    test_analyze(d)
    print("OK")
