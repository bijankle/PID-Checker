"""
End-to-end pipeline test against synthetic data we generate on the fly, so the
core logic is verifiable before real sample files arrive. We build a small PDF
with a title block and some tags, plus a line list, and assert the reconciler
catches the planted inconsistencies.
"""
import os
import sys

import fitz
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pidchecker.pdf_extractor import extract_pdf  # noqa: E402
from pidchecker.excel_loader import load_list      # noqa: E402
from pidchecker.reconciler import reconcile        # noqa: E402


def _make_pdf(path):
    doc = fitz.open()
    page = doc.new_page(width=842, height=595)  # A4 landscape-ish
    # body tags
    page.insert_text((50, 60), '6"-P-1001-A1A appears here')
    page.insert_text((50, 90), 'Valve FV-1001 on this line')
    page.insert_text((50, 120), 'Pump P-101 feeds the system')
    # title block (bottom-right region)
    page.insert_text((620, 540), "DWG NO  PID-100-001")
    doc.save(path)
    doc.close()


def _make_line_list(path):
    df = pd.DataFrame({
        "Line Number": ['6"-P-1001-A1A', '8"-P-2002-B2B'],  # 2nd not on PID
        "Size": ['6"', '8"'],
        "P&ID": ["PID-100-001", "PID-100-001"],
    })
    df.to_excel(path, index=False)


def test_pipeline_detects_inconsistencies(tmp_path):
    pdf = tmp_path / "pid.pdf"
    xls = tmp_path / "line_list.xlsx"
    _make_pdf(str(pdf))
    _make_line_list(str(xls))

    pages = extract_pdf(str(pdf))
    assert pages[0].drawing_number == "PID-100-001"
    assert pages[0].drawing_number_source == "title_block"
    assert '6"-P-1001-A1A'.upper().replace('"', "") in pages[0].tags["line"]

    line_list = load_list(str(xls), "line_list")
    assert line_list.id_column is not None
    assert len(line_list.items) == 2

    findings = reconcile([line_list], pages)
    categories = {f.category for f in findings}
    # 8"-P-2002-B2B is listed but not drawn:
    assert "in_list_not_on_pid" in categories


if __name__ == "__main__":
    import tempfile, pathlib
    d = pathlib.Path(tempfile.mkdtemp())
    test_pipeline_detects_inconsistencies(d)
    print("OK")
