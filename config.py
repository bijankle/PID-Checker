"""
Central configuration for PID-Checker.

Almost everything that depends on YOUR document formats lives here so it can be
calibrated against real sample files without touching the parsing/reconciliation
logic. Once you upload sample P&IDs and lists we tune three things in this file:

  1. TITLE_BLOCK  - where on the page the drawing number lives.
  2. TAG_PATTERNS - the regexes that recognise line/valve/tie-in/equipment tags.
  3. LIST_SCHEMAS - which spreadsheet columns hold the identifier + attributes.
"""

# ---------------------------------------------------------------------------
# 1. Title block / drawing-number extraction
# ---------------------------------------------------------------------------
# The title block is searched first for the drawing (P&ID) number. Region is
# expressed as fractions of page width/height so it works regardless of sheet
# size (A1/A3/ANSI D, etc.). Default = bottom-right quadrant, the usual spot.
TITLE_BLOCK = {
    "x0_frac": 0.55,   # left edge of search region (fraction of page width)
    "y0_frac": 0.70,   # top edge  of search region (fraction of page height)
    "x1_frac": 1.00,   # right edge
    "y1_frac": 1.00,   # bottom edge
}

# Regex used to recognise a drawing / P&ID number. Calibrated to the 2233
# Kathleen Valley drawings, whose P&IDs are numbered AREA-FP-NNN (e.g. 120-FP-001).
DRAWING_NUMBER_PATTERN = r"\d{3}-FP-\d{3}"

# Optional label that precedes the number in the title block, e.g. "DWG NO".
# When present we prefer the token that follows it.
DRAWING_NUMBER_LABELS = ["DWG NO", "DWG. NO", "DRAWING NO", "P&ID NO", "PID NO",
                         "DRAWING NUMBER", "DOCUMENT NO", "SHEET"]

# ---------------------------------------------------------------------------
# 2. Tag patterns found on the drawings
# ---------------------------------------------------------------------------
# These match the identifiers as they appear *on the P&ID*. Calibrate to the
# project's tagging convention. Each is a (name, regex) so report rows are
# labelled by item type.
# Calibrated to the 2233 Kathleen Valley tagging convention. The reconciler
# also uses the MEL's set of two-letter type codes to tell equipment apart from
# other AREA-XX-NNN tags (cameras CA, stockpiles SP, etc.) — see check_project.py.
TAG_PATTERNS = {
    # Line numbers: SIZE-SERV-SPEC-NNN, e.g. 25-PW-S31-027, 100-PA-P35-267
    "line": r"\b\d{2,4}-[A-Z]{2,3}-[A-Z]\d{2}-\d{3}\b",
    # Valves: sized tag 50V11A  OR  pattern tag AREA-VV-NNN (e.g. 120-VV-003)
    "valve": r"\b(?:\d{2,4}V\d{2}[A-Z]|\d{3}-VV-\d{3})\b",
    # Equipment: AREA-TYPE-NNN, e.g. 121-CV-014 (type validated against the MEL)
    "equipment": r"\b\d{3}-[A-Z]{2}-\d{3}\b",
}

# ---------------------------------------------------------------------------
# 3. Spreadsheet schemas
# ---------------------------------------------------------------------------
# For each list type we declare:
#   item_type : which TAG_PATTERNS family it reconciles against
#   id_columns: candidate header names for the unique identifier (first match
#               wins; matched case-insensitively, ignoring spaces/underscores)
#   attributes: candidate header names for attributes we cross-check when the
#               same column also appears on / can be derived from the drawing
# Header matching is fuzzy (normalised), so "Line No", "LINE_NUMBER" and
# "line number" all resolve to the same thing.
LIST_SCHEMAS = {
    "line_list": {
        "item_type": "line",
        "id_columns": ["line number", "line no", "line tag", "line", "tag"],
        "attributes": ["size", "nominal size", "service", "spec", "piping class",
                       "from", "to", "p&id", "pid", "drawing", "sheet"],
    },
    "valve_list": {
        "item_type": "valve",
        "id_columns": ["valve tag", "valve no", "tag no", "tag", "valve"],
        "attributes": ["size", "type", "service", "line number", "line",
                       "p&id", "pid", "drawing", "sheet"],
    },
    "tie_in_list": {
        "item_type": "tie_in",
        "id_columns": ["tie-in", "tie in", "tie-in no", "tie in number",
                       "ti no", "tag", "number"],
        "attributes": ["size", "service", "line number", "existing line",
                       "new line", "p&id", "pid", "drawing", "sheet"],
    },
    "equipment_list": {
        "item_type": "equipment",
        "id_columns": ["equipment tag", "equipment no", "tag no", "item no",
                       "tag", "equipment"],
        "attributes": ["description", "service", "type", "duty",
                       "p&id", "pid", "drawing", "sheet"],
    },
}

# Map an uploaded file (by a keyword in its filename) to a schema above.
# Used to auto-classify uploads; the UI also lets the user override.
FILENAME_HINTS = {
    "line_list": ["line"],
    "valve_list": ["valve"],
    "tie_in_list": ["tie", "tie-in", "tiein"],
    "equipment_list": ["equip", "mechanical"],
}

# ---------------------------------------------------------------------------
# 4. Reconciliation behaviour
# ---------------------------------------------------------------------------
# When a list row names a P&ID/drawing explicitly (via an attribute column),
# also verify the item is actually found on THAT page (not just somewhere).
CHECK_DECLARED_PAGE = True

# Normalisation applied to every identifier before comparison.
def normalize_tag(value: str) -> str:
    """Uppercase, collapse whitespace, and strip surrounding punctuation so that
    '6"-P-1001 ', '6-P-1001' style variations compare consistently."""
    if value is None:
        return ""
    s = str(value).upper().strip()
    s = s.replace('"', "").replace("'", "")
    # collapse internal whitespace
    s = " ".join(s.split())
    return s
