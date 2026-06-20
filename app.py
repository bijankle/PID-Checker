"""
PID-Checker web app.

Drag-and-drop the four project files — P&ID PDF, Line List, Valve List, MEL —
and get the consistency report in the browser plus a downloadable Excel.

The analysis itself lives in check_project.py (shared with the command-line
version) so the web app and the script always agree.
"""
from __future__ import annotations

import os
import uuid

from flask import (Flask, render_template, request, send_file, redirect,
                   url_for, flash)
from werkzeug.utils import secure_filename

from check_project import analyze, write_report

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("PIDCHECKER_SECRET", "dev-secret-change-me")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB total upload

# Remembers the most recent report per run so the Download button works.
_REPORTS: dict[str, str] = {}

REQUIRED = {
    "pid_pdf": "P&ID PDF",
    "line_list": "Line List",
    "valve_list": "Valve List",
    "mel": "Mechanical Equipment List (MEL)",
}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/check", methods=["POST"])
def check():
    # All four files are needed for a full reconciliation.
    missing = [label for field, label in REQUIRED.items()
               if not request.files.get(field) or not request.files[field].filename]
    if missing:
        flash("Please upload: " + ", ".join(missing))
        return redirect(url_for("index"))

    workdir = os.path.join(UPLOAD_DIR, uuid.uuid4().hex)
    os.makedirs(workdir, exist_ok=True)

    paths = {}
    for field in REQUIRED:
        f = request.files[field]
        p = os.path.join(workdir, secure_filename(f.filename))
        f.save(p)
        paths[field] = p

    try:
        df, disc, summary = analyze(
            paths["pid_pdf"], paths["line_list"],
            paths["valve_list"], paths["mel"],
        )
    except Exception as exc:  # noqa: BLE001 - show the problem to the user
        flash(f"Could not analyse the files: {exc}")
        return redirect(url_for("index"))

    run_id = uuid.uuid4().hex
    report_path = os.path.join(workdir, "PID_Check_Report.xlsx")
    write_report(df, disc, summary, report_path)
    _REPORTS[run_id] = report_path

    summary_dict = dict(zip(summary["Metric"], summary["Value"]))
    return render_template(
        "report.html",
        run_id=run_id,
        summary=summary_dict,
        discrepancies=disc.to_dict(orient="records"),
        inventory=df.to_dict(orient="records"),
    )


@app.route("/download/<run_id>")
def download(run_id: str):
    path = _REPORTS.get(run_id)
    if not path or not os.path.exists(path):
        flash("Report no longer available; please re-run the check.")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True,
                     download_name="PID_Check_Report.xlsx")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
