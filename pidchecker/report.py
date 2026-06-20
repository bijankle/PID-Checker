"""Render reconciliation findings to an Excel workbook for download."""
from __future__ import annotations

import pandas as pd

from pidchecker.reconciler import Finding


def findings_to_dataframe(findings: list[Finding]) -> pd.DataFrame:
    rows = [{
        "Severity": f.severity,
        "Category": f.category,
        "Item Type": f.item_type,
        "Identifier": f.identifier,
        "Drawing / Page": f.drawing,
        "Source": f.source,
        "Message": f.message,
    } for f in findings]
    df = pd.DataFrame(rows, columns=[
        "Severity", "Category", "Item Type", "Identifier",
        "Drawing / Page", "Source", "Message",
    ])
    # errors first, then warnings, then info
    order = {"error": 0, "warning": 1, "info": 2}
    if not df.empty:
        df = df.sort_values(
            by=["Severity", "Item Type", "Identifier"],
            key=lambda s: s.map(order) if s.name == "Severity" else s,
        ).reset_index(drop=True)
    return df


def write_excel(findings: list[Finding], summary: dict, path: str) -> None:
    df = findings_to_dataframe(findings)
    summary_rows = [{"Metric": "Errors", "Count": summary.get("error", 0)},
                    {"Metric": "Warnings", "Count": summary.get("warning", 0)},
                    {"Metric": "Info", "Count": summary.get("info", 0)}]
    summary_rows += [{"Metric": f"  {k}", "Count": v}
                     for k, v in summary.get("by_category", {}).items()]
    summary_df = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        summary_df.to_excel(xw, sheet_name="Summary", index=False)
        df.to_excel(xw, sheet_name="Findings", index=False)
