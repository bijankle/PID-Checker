"""
End-to-end test of the calibrated checker on small synthetic files that mimic
the real project's conventions, so the logic is verifiable without shipping the
actual (confidential) drawings.

Covers: tags found on the drawing vs the lists, position-based drawing-number
detection (title block at the bottom, ignoring cross-references), and the
"drawn here but the MEL says a different P&ID" cross-check.
"""
import os
import sys

import fitz
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from check_project import analyze  # noqa: E402


def _make_pdf(path):
    doc = fitz.open()
    page = doc.new_page(width=842, height=595)
    page.insert_text((50, 40), "130-FP-099")           # cross-reference near the top
    page.insert_text((50, 90), "121-CV-014")           # equipment, MEL home = this sheet
    page.insert_text((50, 110), "121-PP-002")          # equipment, MEL home = a DIFFERENT sheet
    page.insert_text((50, 130), "121-CV-099")          # equipment NOT in MEL
    page.insert_text((50, 150), "120-VV-003")          # valve in list
    page.insert_text((50, 170), "50V11A")              # sized valve in list
    page.insert_text((50, 190), "25-PW-S31-027")       # line in list
    page.insert_text((50, 210), "25-PW-S31-999")       # line NOT in list
    page.insert_text((50, 230), "120-CA-202")          # camera: untracked type
    page.insert_text((700, 575), "120-FP-001")         # the real drawing number (title block, bottom-right)
    doc.save(path)
    doc.close()


def _make_lists(tmp_path):
    mel = tmp_path / "mel.xlsx"
    pd.DataFrame({
        "Equipment No.": ["121-CV-014", "121-PP-002"],
        "P&ID\nNo.": ["120-FP-001", "130-FP-002"],   # PP-002's home is a different sheet
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


def test_analyze(tmp_path):
    pdf = tmp_path / "pid.pdf"
    _make_pdf(str(pdf))
    line, valve, mel = _make_lists(tmp_path)

    df, disc, summary = analyze(str(pdf), str(line), str(valve), str(mel))

    # position-based: the bottom title-block number wins over the top cross-ref
    assert set(df["P&ID"]) == {"120-FP-001"}

    status = dict(zip(df["Tag"], df["Status"]))
    assert status["121-CV-099"] == "NOT IN MEL"
    assert status["25-PW-S31-999"] == "NOT IN LINE LIST"
    assert status["120-CA-202"].startswith("NOT IN MEL")
    assert "P&ID MISMATCH" in status["121-PP-002"]   # MEL home is 130-FP-002

    # things that are correct must NOT be flagged
    assert status["121-CV-014"] == "IN LIST"
    assert status["120-VV-003"] == "IN LIST"
    assert status["50V11A"] == "IN LIST"
    assert status["25-PW-S31-027"] == "IN LIST"   # zero-padded list entry still matches


if __name__ == "__main__":
    import tempfile, pathlib
    d = pathlib.Path(tempfile.mkdtemp())
    test_analyze(d)
    print("OK")
