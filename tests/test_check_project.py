"""
End-to-end test of the calibrated checker on small synthetic files that mimic
the real project's conventions, so the logic is verifiable without shipping the
actual (confidential) drawings.
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
    page.insert_text((50, 60), "120-FP-001")          # drawing number
    page.insert_text((50, 90), "121-CV-014")          # equipment in MEL
    page.insert_text((50, 110), "121-CV-099")          # equipment NOT in MEL
    page.insert_text((50, 130), "120-VV-003")         # valve in list
    page.insert_text((50, 150), "50V11A")             # sized valve in list
    page.insert_text((50, 170), "25-PW-S31-027")      # line in list
    page.insert_text((50, 190), "25-PW-S31-999")      # line NOT in list
    page.insert_text((50, 210), "120-CA-202")         # camera: untracked type
    doc.save(path)
    doc.close()


def _make_lists(tmp_path):
    mel = tmp_path / "mel.xlsx"
    pd.DataFrame({"Equipment No.": ["121-CV-014", "121-PP-002"]}).to_excel(
        mel, sheet_name="Equipment List", index=False)

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
    # row 0 is the header, row 1 is a sub-header (skipped by .iloc[1:]), then data
    df = pd.DataFrame([
        ["sub", "", "", ""],
        ["", "PW", "S31", "27"],
    ], columns=["Identifier", "SERV CODE", "SPEC", "LINE No."])
    df.to_excel(line, sheet_name="Line List", index=False)
    return line, valve, mel


def test_analyze_flags_only_pid_not_in_list(tmp_path):
    pdf = tmp_path / "pid.pdf"
    _make_pdf(str(pdf))
    line, valve, mel = _make_lists(tmp_path)

    df, disc, summary = analyze(str(pdf), str(line), str(valve), str(mel))

    flagged = set(disc["Tag"])
    assert "121-CV-099" in flagged       # equipment not in MEL
    assert "25-PW-S31-999" in flagged    # line not in line list
    assert "120-CA-202" in flagged       # untracked camera -> discrepancy

    # things that ARE in the lists must not be flagged
    assert "121-CV-014" not in flagged
    assert "120-VV-003" not in flagged
    assert "50V11A" not in flagged
    assert "25-PW-S31-027" not in flagged


if __name__ == "__main__":
    import tempfile, pathlib
    d = pathlib.Path(tempfile.mkdtemp())
    test_analyze_flags_only_pid_not_in_list(d)
    print("OK")
