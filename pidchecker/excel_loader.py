"""
Load engineering lists (line/valve/tie-in/equipment) from .xls/.xlsx and
normalise them into a common shape so the reconciler is format-agnostic.

Header rows in real lists are often not the first row (there's a title banner,
revision block, etc.), so we search the first several rows for the one that best
matches the expected columns for the chosen schema.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

import config


@dataclass
class ListItem:
    identifier: str                       # normalised unique tag
    raw_identifier: str                   # as written in the sheet
    row_number: int                       # 1-based row in the source file
    attributes: dict = field(default_factory=dict)
    declared_drawing: Optional[str] = None  # P&ID named on the row, if any


@dataclass
class LoadedList:
    schema_name: str                      # e.g. "line_list"
    item_type: str                        # e.g. "line"
    source_file: str
    items: list[ListItem] = field(default_factory=list)
    id_column: Optional[str] = None       # resolved header used as identifier
    warnings: list[str] = field(default_factory=list)


def _norm_header(name) -> str:
    return "".join(str(name).lower().split()).replace("_", "").replace("-", "")


def _resolve_column(df_columns, candidates) -> Optional[str]:
    """Return the actual df column matching any candidate (fuzzy, normalised)."""
    norm_map = {_norm_header(c): c for c in df_columns}
    for cand in candidates:
        key = _norm_header(cand)
        if key in norm_map:
            return norm_map[key]
    # also allow substring matches (e.g. "linenumber" inside "processlinenumber")
    for cand in candidates:
        key = _norm_header(cand)
        for ncol, col in norm_map.items():
            if key and key in ncol:
                return col
    return None


def _read_with_header_detection(path: str, candidates, max_scan: int = 8):
    """Read the sheet, locating the most likely header row within the first
    `max_scan` rows by counting how many candidate columns it exposes."""
    engine = None  # let pandas pick; openpyxl for xlsx, xlrd for legacy xls
    raw = pd.read_excel(path, header=None, dtype=str, engine=engine)
    best_row, best_score = 0, -1
    for r in range(min(max_scan, len(raw))):
        header_vals = raw.iloc[r].tolist()
        if _resolve_column(header_vals, candidates):
            score = sum(
                1 for c in raw.iloc[r].tolist()
                if str(c).strip() and str(c).lower() != "nan"
            )
            if score > best_score:
                best_row, best_score = r, score
    df = pd.read_excel(path, header=best_row, dtype=str, engine=engine)
    df = df.dropna(how="all")
    return df, best_row


def load_list(path: str, schema_name: str) -> LoadedList:
    schema = config.LIST_SCHEMAS[schema_name]
    result = LoadedList(
        schema_name=schema_name,
        item_type=schema["item_type"],
        source_file=path,
    )

    df, header_row = _read_with_header_detection(path, schema["id_columns"])
    id_col = _resolve_column(df.columns, schema["id_columns"])
    if id_col is None:
        result.warnings.append(
            f"Could not find an identifier column among {schema['id_columns']}. "
            f"Columns present: {list(df.columns)}"
        )
        return result
    result.id_column = id_col

    attr_cols = {}
    for attr in schema["attributes"]:
        col = _resolve_column(df.columns, [attr])
        if col is not None:
            attr_cols[attr] = col

    drawing_col = None
    for key in ("p&id", "pid", "drawing", "sheet"):
        if key in attr_cols:
            drawing_col = attr_cols[key]
            break

    for offset, (_, row) in enumerate(df.iterrows()):
        raw_id = row.get(id_col)
        if raw_id is None or str(raw_id).strip() == "" or str(raw_id).lower() == "nan":
            continue
        ident = config.normalize_tag(raw_id)
        attrs = {
            a: (None if str(row.get(c)).lower() == "nan" else str(row.get(c)).strip())
            for a, c in attr_cols.items()
        }
        declared = None
        if drawing_col is not None:
            v = row.get(drawing_col)
            if v is not None and str(v).strip() and str(v).lower() != "nan":
                declared = config.normalize_tag(v)
        result.items.append(
            ListItem(
                identifier=ident,
                raw_identifier=str(raw_id).strip(),
                row_number=header_row + 2 + offset,  # +2: header + 1-based
                attributes=attrs,
                declared_drawing=declared,
            )
        )

    return result
