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

# Regex used to recognise a drawing / P&ID number. This is intentionally broad;
# narrow it once we see the real numbering scheme (e.g. r"PID-\d{3}-\d{4}").
DRAWING_NUMBER_PATTERN = r"[A-Z0-9]{1,5}[-_][A-Z0-9]{2,6}(?:[-_][A-Z0-9]{1,6}){0,3}"

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
TAG_PATTERNS = {
    # e.g. 6"-P-1001-A1A  /  6-P-1001  /  150-CWS-001
    "line": r'\b\d{1,2}"?-?[A-Z]{1,4}-?\d{3,5}(?:-[A-Z0-9]{1,5})?\b',
    # e.g. FV-1001, HV-200A, PSV-3001
    "valve": r"\b[A-Z]{1,4}V-\d{2,5}[A-Z]?\b",
    # e.g. TI-12, TIE-045, T-001
    "tie_in": r"\bTIE?[-_]?\d{1,4}[A-Z]?\b",
    # e.g. P-101, E-2003, V-12A, C-300
    "equipment": r"\b[A-Z]{1,3}-\d{2,4}[A-Z]?\b",
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
