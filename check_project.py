"""
Calibrated P&ID-vs-list reconciliation for the 2233 Kathleen Valley project
(Lycopodium drawings + Line List / Valve List / MEL).

Output is organised as one block per item type, with the value read from the
P&ID beside the value from the list, so disagreements can be eyeballed (and are
red-filled in the Excel report):

  Equipment : PID equip-no | MEL equip-no | PID description | MEL description
  Valves    : PID valve tag | list valve tag | PID size-code | list size-code
  Lines     : PID line-no  | list line-no  | list P&ID      | sheet(s) found on

Conventions confirmed against the sample files:
  Drawing number : title block, AREA-FP-NNN (read by position: nearest the
                   bottom of the sheet, so cross-references aren't mistaken for it)
  Equipment tags : AREA-TYPE-NNN; must appear >=2 times on a sheet to count as
                   belonging to it (genuine equipment is labelled at the symbol
                   and in the reference block; cross-references appear once)
  Valve tags     : sized <DN>V<nn><L> (e.g. 25V41A) AND unique AREA-VV-NNN
                   (e.g. 120-VV-008); the two are paired on the drawing by proximity
  Line tags      : <SIZE>-<SERV>-<SPEC>-<NNN> (e.g. 25-PW-S31-027)

Run:  python check_project.py PID.pdf Line_List.xlsx Valve_List.xlsx MEL.xlsx [out.xlsx]
"""
from __future__ import annotations

import re
import sys
from collections import Counter

import fitz
import pandas as pd

# --- regexes for the on-drawing tag conventions ---------------------------
RE_EQ   = re.compile(r"\b(\d{3})-([A-Z]{2})-(\d{3})\b")
RE_VV   = re.compile(r"\b(\d{3})-VV-(\d{3})\b")
RE_VSZ  = re.compile(r"\b(\d{2,4})V(\d{2})([A-Z])\b")
RE_LINE = re.compile(r"\b(\d{2,4})-([A-Z]{2,3})-([A-Z]\d{2})-(\d{3})\b")
RE_DWG  = re.compile(r"\b(\d{3}-FP-\d{3})\b")
RE_FP_FULL  = re.compile(r"^\d{3}-FP-\d{3}$")
RE_VSZ_FULL = re.compile(r"^\d{2,4}V\d{2}[A-Z]$")
RE_VV_FULL  = re.compile(r"^\d{3}-VV-\d{3}$")

NON_EQUIPMENT_CODES = {"FP"}   # drawing references; VV handled separately
ONLY_ONCE = "Only labelled once — verify (possible cross-reference)"


def _find_col(columns, *wanted):
    norm = lambda s: "".join(str(s).lower().split()).replace("_", "").replace("-", "")
    targets = [norm(w) for w in wanted]
    for c in columns:
        if norm(c) in targets:
            return c
    return None


def _norm_desc(s):
    """Loose normalisation for comparing free-text descriptions."""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def descriptions_match(pid_desc, mel_desc):
    a, b = _norm_desc(pid_desc), _norm_desc(mel_desc)
    if not a or not b:
        return False
    return a == b or a in b or b in a


# --------------------------------------------------------------------------
# page view: words + lines + text, from the native text layer, falling back to
# OCR for sheets that have no extractable text (scanned, or text-as-outlines).
# --------------------------------------------------------------------------
def page_view(page):
    """Words + lines + text from the native text layer. Sheets with no
    extractable text (scanned, or text saved as vector outlines) come back
    empty and are surfaced for manual review rather than OCR'd."""
    lines = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            txt = " ".join(s["text"] for s in l["spans"]).strip()
            if txt:
                x0, y0, x1, y1 = l["bbox"]
                lines.append((x0, y0, x1, y1, txt))
    return {"W": page.rect.width, "H": page.rect.height,
            "words": [tuple(w[:5]) for w in page.get_text("words")],
            "lines": lines, "text": page.get_text("text")}


# --------------------------------------------------------------------------
# positional extraction from a page view
# --------------------------------------------------------------------------
def drawing_number(pv):
    """Sheet's own number = the FP word in the bottom-right title block. Prefer
    candidates in the bottom-right corner; fall back to the lowest FP anywhere."""
    fps = [w for w in pv["words"] if RE_FP_FULL.match(w[4])]
    if not fps:
        return None
    W, H = pv["W"], pv["H"]
    corner = [w for w in fps if w[0] > 0.6 * W and w[1] > 0.85 * H]
    pool = corner or fps
    pool.sort(key=lambda w: (w[1], w[0]))   # lowest, then right-most
    return pool[-1][4]


def _is_tagish(t):
    t = t.strip().upper()
    return bool(RE_EQ.fullmatch(t) or re.match(r"^\d{2,4}-[A-Z]{2,3}-[A-Z0-9]", t)
                or re.fullmatch(r"\d{1,4}", t) or re.fullmatch(r"[A-Z]{1,5}\s?\d{2,4}", t))


def _is_noise_line(t):
    """Lines that are not part of an equipment description (drawing/vendor refs)."""
    u = t.strip().upper()
    return bool(re.search(r"-DRG-|-ME-|\bVENDOR\b|\bPACKAGE\b", u)
                or re.match(r"^\d{3,4}-\d", u))


def _desc_reliable(d):
    """A description we trust enough to red-flag a mismatch on: has real words,
    isn't a stray code/ref."""
    if not d or _is_noise_line(d):
        return False
    return len(re.findall(r"[A-Z]{3,}", d.upper())) >= 2


def equipment_descriptions(pv):
    """tag -> description, read from the reference block (the 1-3 tightly stacked
    lines directly beneath the tag at the same x). The symbol occurrence has no
    such block, so the longest clean description found across occurrences wins."""
    H = pv["H"]
    lines = pv["lines"]
    out = {}
    for s in lines:
        tag = s[4].strip().upper()
        m = RE_EQ.fullmatch(tag)
        if not m or m.group(2) in NON_EQUIPMENT_CODES or m.group(2) == "VV":
            continue
        ex0, ey0 = s[0], s[1]
        desc, cur = [], s[3]
        for t in sorted([l for l in lines if l[1] > ey0], key=lambda z: z[1]):
            if abs(t[0] - ex0) > 40 or t[1] - cur > 0.02 * H or t[1] < cur - 1:
                continue
            if _is_tagish(t[4]) or _is_noise_line(t[4]):
                break          # stop at the next tag or a vendor/drawing-ref line
            desc.append(t[4].strip())
            cur = t[3]
            if len(desc) >= 3:
                break
        d = " ".join(desc)
        if len(d) > len(out.get(tag, "")):
            out[tag] = d
    return out


def valve_pairs(pv):
    """unique valve tag (AREA-VV-NNN) -> nearest sized code (<DN>V<nn><L>) on the
    sheet, when within ~5% of sheet width (they sit together at the valve symbol)."""
    ws = pv["words"]
    W = pv["W"]
    sized = [(w[4], (w[0] + w[2]) / 2, (w[1] + w[3]) / 2) for w in ws if RE_VSZ_FULL.match(w[4])]
    out = {}
    for w in ws:
        if not RE_VV_FULL.match(w[4]):
            continue
        vx, vy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
        best, bd = "", 1e9
        for stag, sx, sy in sized:
            dd = ((vx - sx) ** 2 + (vy - sy) ** 2) ** 0.5
            if dd < bd:
                bd, best = dd, stag
        if best and bd < 0.05 * W:
            out[w[4]] = best
    return out


def continuation_ribbons(pv):
    """Off-page connectors sit at the left/right border: the target drawing
    number in the flag, with the line number on the same row just inboard.
    Returns a list of (target_drawing, line_tag) for this sheet."""
    W, H = pv["W"], pv["H"]
    ws = pv["words"]
    fps = [w for w in ws if RE_FP_FULL.match(w[4])]
    lns = [w for w in ws if RE_LINE.fullmatch(w[4])]
    out = []
    for w in fps:
        xf, yf = w[0] / W, w[1] / H
        if yf >= 0.88 or (0.12 < xf < 0.88):
            continue                       # title block, or not a border connector
        fx, fy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
        cands = [l for l in lns if abs((l[1] + l[3]) / 2 - fy) < 0.015 * H]  # same row
        if not cands:
            continue
        best = min(cands, key=lambda l: abs((l[0] + l[2]) / 2 - fx))
        if abs((best[0] + best[2]) / 2 - fx) / W < 0.15:
            out.append((w[4], best[4]))
    return out


# --------------------------------------------------------------------------
# reference data from the lists
# --------------------------------------------------------------------------
def load_references(line_path, valve_path, mel_path):
    mel = pd.read_excel(mel_path, sheet_name="Equipment List", header=0, dtype=str)
    equipment, codes, declared_pid, names = set(), set(), {}, {}
    eq_col = _find_col(mel.columns, "Equipment No.") or "Equipment No."
    pid_col = _find_col(mel.columns, "P&ID No.", "P&ID\nNo.", "PID No.")
    name_col = _find_col(mel.columns, "Equipment Name", "Description")
    for _, row in mel.iterrows():
        m = RE_EQ.fullmatch(str(row.get(eq_col)).strip().upper())
        if not m:
            continue
        equipment.add(m.group(0))
        codes.add(m.group(2))
        if pid_col is not None:
            dm = RE_DWG.search(str(row.get(pid_col)).upper())
            if dm:
                declared_pid[m.group(0)] = dm.group(1)
        if name_col is not None:
            v = str(row.get(name_col)).strip()
            if v and v.lower() != "nan":
                names[m.group(0)] = v

    # Valves: the unique VV tag and its paired size code, plus the full id set.
    valves, vv_set, vv_to_size = set(), set(), {}
    vmain = pd.read_excel(valve_path, sheet_name="2233-PLST-007", header=0, dtype=str)
    patt_col = _find_col(vmain.columns, "PATTERN")
    name_lk = _find_col(vmain.columns, "Valve Name Lookup")
    for col in ("Valve Identifier", "Valve Name Lookup", "PATTERN"):
        if col in vmain.columns:
            valves |= {str(v).strip().upper() for v in vmain[col].dropna()}
    for _, r in vmain.iterrows():
        patt = str(r.get(patt_col)).strip().upper() if patt_col else ""
        if RE_VV_FULL.match(patt):
            vv_set.add(patt)
            nm = str(r.get(name_lk)).strip().upper() if name_lk else ""
            if RE_VSZ_FULL.match(nm):
                vv_to_size[patt] = nm
    for sheet, col in (("Actuated Valves", "TAG"), ("Actuated Valves", "Valve Code"),
                       ("Manual Valves", "Valve Tag")):
        try:
            s = pd.read_excel(valve_path, sheet_name=sheet, header=0, dtype=str)
            valves |= {str(v).strip().upper() for v in s[col].dropna()}
        except Exception:
            pass
    valves = {v for v in valves if v and v not in {"NAN", "-", "SPARE"}}

    # Lines are identified by SERVICE + SEQUENTIAL only; size and spec may change
    # along a line (reducers / spec breaks) and it is still the same line. We key
    # everything on (service, sequential) and record which P&ID(s) the list
    # assigns each line to, plus the list's own line identifier.
    ll = pd.read_excel(line_path, sheet_name="Line List", header=0, dtype=str).iloc[1:]
    serv_num, line_pid, line_ident = set(), {}, {}
    pidc = _find_col(ll.columns, "P&ID", "PID")
    ident_col = _find_col(ll.columns, "Identifier")

    for _, r in ll.iterrows():
        pids = set(RE_DWG.findall(str(r.get(pidc)).upper())) if pidc else set()
        serv = str(r.get("SERV CODE")).strip().upper()
        num = re.sub(r"\D", "", str(r.get("LINE No.")))
        if serv and num:
            key = (serv, str(int(num)))
            serv_num.add(key)
            line_pid.setdefault(key, set()).update(pids)
            ident = str(r.get(ident_col)).strip() if ident_col else ""
            if ident and ident.lower() != "nan":
                line_ident.setdefault(key, ident)
        for c in ll.columns:   # harvest fully-written line numbers in any cell
            for m in RE_LINE.finditer(str(r.get(c)).upper()):
                k = (m.group(2), str(int(m.group(4))))
                serv_num.add(k)
                line_pid.setdefault(k, set())

    return {"equipment": equipment, "codes": codes, "valves": valves,
            "vv_set": vv_set, "vv_to_size": vv_to_size, "serv_num": serv_num,
            "line_pid": line_pid, "line_ident": line_ident,
            "declared_pid": declared_pid, "names": names}


def page_tags(text, codes):
    T = text.upper()
    lines = {m.group(0) for m in RE_LINE.finditer(T)}
    valves, vvs = set(), set()
    eq_counts, other_counts = Counter(), Counter()
    for m in RE_EQ.finditer(T):
        code, tag = m.group(2), m.group(0)
        if code in NON_EQUIPMENT_CODES:
            continue
        if code == "VV":
            vvs.add(tag)
        elif code in codes:
            eq_counts[tag] += 1
        else:
            other_counts[tag] += 1
    for m in RE_VSZ.finditer(T):
        valves.add(m.group(0))
    return valves, vvs, lines, eq_counts, other_counts


def line_key(tag):
    """A line's identity = (service, sequential)."""
    m = RE_LINE.fullmatch(tag)
    return (m.group(2), str(int(m.group(4))))


def line_in_list(tag, ref):
    return line_key(tag) in ref["serv_num"]


# --------------------------------------------------------------------------
# main analysis -> one table (with a red-fill mask) per item type
# --------------------------------------------------------------------------
def analyze(pdf_path, line_path, valve_path, mel_path):
    ref = load_references(line_path, valve_path, mel_path)
    declared, names = ref["declared_pid"], ref["names"]

    # ---- pass 1: read every page ----
    pages, homes, manual_pages = [], {}, []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            pv = page_view(page)
            if len(pv["text"].strip()) < 20:        # no readable text -> manual check
                manual_pages.append(page.number + 1)
                continue
            dwg = drawing_number(pv) or f"(page {page.number + 1})"
            valves, vvs, lines, eq_counts, other_counts = page_tags(pv["text"], ref["codes"])
            pages.append({
                "dwg": dwg, "vvs": vvs, "lines": lines,
                "eq": eq_counts, "other": other_counts,
                "desc": equipment_descriptions(pv),
                "vpairs": valve_pairs(pv),
                "ribbons": continuation_ribbons(pv),
            })
            for tag, c in (eq_counts | other_counts).items():
                if c >= 2:
                    homes.setdefault(tag, set()).add(dwg)

    eq_rows, eq_mask = [], []
    vl_rows, vl_mask = [], []
    ln_rows, ln_mask = [], []
    ot_rows = []

    # ---- pass 2 ----
    for pg in pages:
        dwg = pg["dwg"]

        # Equipment: PID no | MEL no | PID descr | MEL descr | Notes
        for tag in sorted(pg["eq"]):
            note = ""
            if pg["eq"][tag] < 2:
                if tag in homes:
                    continue
                note = ONLY_ONCE
            in_mel = tag in ref["equipment"]
            if not in_mel:
                note = (note + "; " if note else "") + "not in MEL"
            elif tag in declared and declared[tag] != dwg:
                note = (note + "; " if note else "") + f"MEL P&ID = {declared[tag]}"
            pid_desc = pg["desc"].get(tag, "")
            mel_desc = names.get(tag, "")
            desc_bad = (in_mel and _desc_reliable(pid_desc) and bool(mel_desc)
                        and not descriptions_match(pid_desc, mel_desc))
            eq_rows.append([dwg, tag, tag if in_mel else "", pid_desc, mel_desc, note])
            eq_mask.append([False, False, not in_mel, desc_bad, desc_bad, bool(note)])

        for tag in sorted(pg["other"]):
            note = ""
            if pg["other"][tag] < 2:
                if tag in homes:
                    continue
                note = ONLY_ONCE
            ot_rows.append([dwg, tag, "NOT IN MEL (untracked type)", note])

        # Valves: PID tag | list tag | PID size | list size | Notes
        for vv in sorted(pg["vvs"]):
            in_list = vv in ref["vv_set"] or vv in ref["valves"]
            pid_size = pg["vpairs"].get(vv, "")
            list_size = ref["vv_to_size"].get(vv, "")
            size_bad = bool(pid_size) and bool(list_size) and pid_size != list_size
            note = "" if in_list else "not in valve list"
            vl_rows.append([dwg, vv, vv if in_list else "", pid_size, list_size, note])
            vl_mask.append([False, False, not in_list, size_bad, size_bad, bool(note)])

    # Lines are grouped across the whole document by (service, sequential), so a
    # continuation line spanning two sheets is one row capturing both P&IDs.
    line_groups = {}   # (serv, seq) -> {"tags": set, "pids": set}
    for pg in pages:
        for tag in pg["lines"]:
            g = line_groups.setdefault(line_key(tag), {"tags": set(), "pids": set()})
            g["tags"].add(tag)
            g["pids"].add(pg["dwg"])
    for key in sorted(line_groups):
        g = line_groups[key]
        matched = key in ref["serv_num"]
        list_no = ref["line_ident"].get(key, "") if matched else ""
        if matched and not list_no:
            list_no = key[0] + key[1]
        pid_pids = sorted(g["pids"])
        list_pids = sorted(ref["line_pid"].get(key, set()))
        # red only if the line is drawn on a sheet the line-list P&ID doesn't include
        pids_bad = matched and bool(list_pids) and not set(pid_pids).issubset(set(list_pids))
        note = "not in line list" if not matched else (
            "drawn on a P&ID not in the line list" if pids_bad else "")
        ln_rows.append([" / ".join(sorted(g["tags"])), list_no,
                        ", ".join(pid_pids), ", ".join(list_pids), note])
        ln_mask.append([False, not matched, pids_bad, pids_bad, bool(note)])

    # Continuations: each off-page ribbon (on a sheet, naming a target drawing
    # and a line) is checked strictly — the named drawing must carry that exact
    # line (fluid service + sequential). Line No (to) is the line as found on the
    # target; a line not on the named drawing is flagged (wrong P&ID / renumber).
    all_dwgs = {pg["dwg"] for pg in pages}
    lines_on = {}
    for pg in pages:
        d = lines_on.setdefault(pg["dwg"], {})
        for tag in pg["lines"]:
            d.setdefault(line_key(tag), tag)
    ct_rows, ct_mask = [], []
    for pg in pages:
        for tgt, ltag in pg["ribbons"]:
            key = line_key(ltag)
            if tgt not in all_dwgs:
                ct_rows.append([ltag, "", pg["dwg"], tgt, "target drawing not in uploaded set"])
                ct_mask.append([False, False, False, False, False])
            elif key in lines_on.get(tgt, {}):
                ct_rows.append([ltag, lines_on[tgt][key], pg["dwg"], tgt, "OK"])
                ct_mask.append([False, False, False, False, False])
            else:
                ct_rows.append([ltag, "", pg["dwg"], tgt, "line not found on named P&ID"])
                ct_mask.append([False, True, False, True, True])

    mr_rows = [[f"page {n}", "No readable text on this sheet — check the P&ID number "
                "and contents manually"] for n in manual_pages]

    eq_cols = ["P&ID", "PID Equip No", "MEL Equip No", "PID Description", "MEL Description", "Notes"]
    vl_cols = ["P&ID", "PID Valve Tag", "List Valve Tag", "PID Size Code", "List Size Code", "Notes"]
    ln_cols = ["PID Line No", "List Line No", "PID P&IDs", "List P&ID", "Notes"]
    ot_cols = ["P&ID", "Tag", "Status", "Notes"]
    ct_cols = ["Line No (from)", "Line No (to)", "P&ID (from)", "P&ID (to)", "Status"]
    mr_cols = ["Sheet", "Issue"]

    tables = {
        "Equipment": pd.DataFrame(eq_rows, columns=eq_cols),
        "Valves":    pd.DataFrame(vl_rows, columns=vl_cols),
        "Lines":     pd.DataFrame(ln_rows, columns=ln_cols),
        "Continuations": pd.DataFrame(ct_rows, columns=ct_cols),
        "Other":     pd.DataFrame(ot_rows, columns=ot_cols),
        "Manual Review": pd.DataFrame(mr_rows, columns=mr_cols),
    }
    masks = {
        "Equipment": pd.DataFrame(eq_mask, columns=eq_cols),
        "Valves":    pd.DataFrame(vl_mask, columns=vl_cols),
        "Lines":     pd.DataFrame(ln_mask, columns=ln_cols),
        "Continuations": pd.DataFrame(ct_mask, columns=ct_cols),
        "Other":     pd.DataFrame([[False] * len(ot_cols) for _ in ot_rows], columns=ot_cols),
        "Manual Review": pd.DataFrame([[True, True] for _ in mr_rows], columns=mr_cols),
    }
    bad_ct = int(pd.DataFrame(ct_mask).any(axis=1).sum()) if ct_mask else 0
    flagged = sum(int(m.any(axis=1).sum()) for m in masks.values()) + len(ot_rows)
    summary = pd.DataFrame({
        "Metric": ["P&ID pages", "Equipment found", "Valves found", "Lines found",
                   "Continuations checked", "Bad continuations",
                   "Other tagged items", "Sheets needing manual check",
                   "Rows needing review"],
        "Value": [len({p["dwg"] for p in pages}), len(eq_rows), len(vl_rows),
                  len(ln_rows), len(ct_rows), bad_ct,
                  len(ot_rows), len(mr_rows), flagged],
    })
    return {"tables": tables, "masks": masks, "summary": summary}


PID_LIST_COLS = {"PID P&IDs", "List P&ID"}


def _col_width(header, cells):
    """Tag/code/P&ID columns hug their longest token; free-text columns size to
    ~their 90th-percentile line length (capped) so they stay readable without
    big blank space."""
    segs, tokenish = [], True
    for v in cells:                       # judge tokenish from values, not the header
        s = str(v) if v is not None else ""
        for seg in s.split("\n"):         # each stacked line is its own segment
            if " " in seg.strip():
                tokenish = False
            segs.append(len(seg))
    if not segs:
        segs = [0]
    segs.sort()
    hlen = len(str(header))
    maxlen = segs[-1]
    p90 = segs[min(len(segs) - 1, int(len(segs) * 0.9))]
    if tokenish:
        return min(max(maxlen, hlen) + 2, 16)
    return min(max(p90, hlen, 18) + 1, 46)


def write_report(result, out_path):
    """One sheet per item type: themed header, red-filled mismatches, wrap text,
    frozen + filtered top row, P&IDs stacked on separate lines, tuned widths."""
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    red_fill = PatternFill("solid", fgColor="FFFCEDED")
    red_font = Font(color="FFA82C2E", bold=True)
    head_fill = PatternFill("solid", fgColor="FF5865F2")
    head_font = Font(color="FFFFFFFF", bold=True)
    thin = Side(style="thin", color="FFE3E5E8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    sheets = [("Summary", result["summary"], None)]
    sheets += [(n, result["tables"][n], result["masks"][n])
               for n in ("Equipment", "Valves", "Lines", "Continuations", "Other", "Manual Review")]
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        for name, df, mask in sheets:
            disp = df.copy()
            for c in disp.columns:
                if c in PID_LIST_COLS:
                    disp[c] = disp[c].map(lambda v: str(v).replace(", ", "\n"))
            disp.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ncols = len(disp.columns)
            for ci, col in enumerate(disp.columns, start=1):
                hc = ws.cell(row=1, column=ci)
                hc.fill, hc.font, hc.border, hc.alignment = head_fill, head_font, border, wrap
            for ri in range(len(disp)):
                for ci in range(ncols):
                    cell = ws.cell(row=ri + 2, column=ci + 1)
                    cell.border, cell.alignment = border, wrap
                    if mask is not None and bool(mask.iat[ri, ci]):
                        cell.fill, cell.font = red_fill, red_font
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{len(disp) + 1}"
            for ci, col in enumerate(disp.columns, start=1):
                ws.column_dimensions[get_column_letter(ci)].width = _col_width(col, disp[col].tolist())


def check(pdf_path, line_path, valve_path, mel_path, out_path="PID_Check_Report.xlsx"):
    result = analyze(pdf_path, line_path, valve_path, mel_path)
    write_report(result, out_path)
    t = result["tables"]
    print(f"Wrote {out_path}: {len(t['Equipment'])} equipment, "
          f"{len(t['Valves'])} valves, {len(t['Lines'])} lines")
    return result


if __name__ == "__main__":
    args = sys.argv[1:]
    if len(args) < 4:
        print(__doc__)
        sys.exit(1)
    check(*args)
