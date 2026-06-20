"""
PID-Checker web app.

Upload a P&ID PDF plus any of the engineering lists (line / valve / tie-in /
equipment). The app extracts each page's drawing number and tags, reconciles the
lists against the drawings, and shows an inconsistency report with a downloadable
Excel version.
"""
from __future__ import annotations

import os
import tempfile
import uuid

from flask import (Flask, render_template, request, send_file, redirect,
                   url_for, flash)
from werkzeug.utils import secure_filename

import config
from pidchecker.pdf_extractor import extract_pdf
from pidchecker.excel_loader import load_list
from pidchecker.reconciler import reconcile, summarize
from pidchecker.report import write_excel, findings_to_dataframe

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("PIDCHECKER_SECRET", "dev-secret-change-me")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB total upload

# In-memory store of generated reports (path) keyed by run id. Fine for a
# single-process local tool; swap for a store if this is ever scaled out.
_REPORTS: dict[str, str] = {}


def _classify(filename: str) -> str | None:
    name = filename.lower()
    for schema, hints in config.FILENAME_HINTS.items():
        if any(h in name for h in hints):
            return schema
    return None


@app.route("/")
def index():
    return render_template("index.html", schemas=config.LIST_SCHEMAS)


@app.route("/check", methods=["POST"])
def check():
    pdf = request.files.get("pid_pdf")
    if not pdf or not pdf.filename:
        flash("Please upload a P&ID PDF.")
        return redirect(url_for("index"))

    workdir = os.path.join(UPLOAD_DIR, uuid.uuid4().hex)
    os.makedirs(workdir, exist_ok=True)

    pdf_path = os.path.join(workdir, secure_filename(pdf.filename))
    pdf.save(pdf_path)
    pages = extract_pdf(pdf_path)

    lists = []
    for field_name in ("line_list", "valve_list", "tie_in_list", "equipment_list"):
        f = request.files.get(field_name)
        if not f or not f.filename:
            continue
        # explicit field overrides filename-based classification
        schema = field_name if field_name in config.LIST_SCHEMAS else _classify(f.filename)
        if not schema:
            continue
        list_path = os.path.join(workdir, secure_filename(f.filename))
        f.save(list_path)
        try:
            lists.append(load_list(list_path, schema))
        except Exception as exc:  # noqa: BLE001 - surface load errors in report
            flash(f"Could not read {f.filename}: {exc}")

    findings = reconcile(lists, pages)
    summary = summarize(findings)

    run_id = uuid.uuid4().hex
    report_path = os.path.join(workdir, "pid_check_report.xlsx")
    write_excel(findings, summary, report_path)
    _REPORTS[run_id] = report_path

    df = findings_to_dataframe(findings)
    return render_template(
        "report.html",
        run_id=run_id,
        summary=summary,
        findings=df.to_dict(orient="records"),
        pages=pages,
        lists=lists,
    )


@app.route("/download/<run_id>")
def download(run_id: str):
    path = _REPORTS.get(run_id)
    if not path or not os.path.exists(path):
        flash("Report no longer available; please re-run the check.")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True,
                     download_name="pid_check_report.xlsx")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
