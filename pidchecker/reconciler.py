"""
Cross-check loaded lists against what was extracted from the P&IDs and produce a
flat list of findings. Each finding has a severity, a category, the item, and
the drawing context so the report can point straight at the problem.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import config
from pidchecker.excel_loader import LoadedList
from pidchecker.pdf_extractor import PageResult


@dataclass
class Finding:
    severity: str          # "error" | "warning" | "info"
    category: str          # short machine-ish label
    item_type: str         # line / valve / tie_in / equipment
    identifier: str        # the tag involved
    message: str           # human-readable explanation
    source: str = ""       # list file or "P&ID"
    drawing: str = ""      # relevant P&ID number / page


@dataclass
class DrawingIndex:
    """Inverted view of the PDF extraction: tag -> drawings it appears on."""
    by_type: dict = field(default_factory=dict)   # item_type -> {tag -> {dwg}}
    pages: list = field(default_factory=list)      # raw PageResult list
    number_to_pages: dict = field(default_factory=dict)  # dwg -> [page_index]

    @classmethod
    def build(cls, pages: list[PageResult]) -> "DrawingIndex":
        by_type: dict[str, dict[str, set]] = {}
        number_to_pages: dict[str, list] = {}
        for p in pages:
            dwg = p.drawing_number or f"(page {p.page_index + 1}, no number)"
            number_to_pages.setdefault(dwg, []).append(p.page_index)
            for item_type, tags in p.tags.items():
                bucket = by_type.setdefault(item_type, {})
                for t in tags:
                    bucket.setdefault(t, set()).add(dwg)
        return cls(by_type=by_type, pages=pages, number_to_pages=number_to_pages)

    def drawings_for(self, item_type: str, tag: str) -> set:
        return self.by_type.get(item_type, {}).get(tag, set())


def reconcile(lists: list[LoadedList], pages: list[PageResult]) -> list[Finding]:
    index = DrawingIndex.build(pages)
    findings: list[Finding] = []

    # Surface pages where no drawing number could be determined.
    for p in pages:
        if not p.drawing_number:
            findings.append(Finding(
                severity="warning",
                category="missing_drawing_number",
                item_type="drawing",
                identifier=f"page {p.page_index + 1}",
                message=("No P&ID number found in the title block or page text. "
                         "Items on this page cannot be attributed to a drawing."),
                source="P&ID",
                drawing=f"page {p.page_index + 1}",
            ))
        elif p.drawing_number_source != "title_block":
            findings.append(Finding(
                severity="info",
                category="drawing_number_outside_title_block",
                item_type="drawing",
                identifier=p.drawing_number,
                message=("Drawing number was found by scanning the page, not in "
                         "the title-block region — verify it is correct."),
                source="P&ID",
                drawing=p.drawing_number,
            ))

    # Track which drawing tags were matched by some list (to flag the reverse).
    matched_on_drawings: dict[str, set] = {}

    for lst in lists:
        for w in lst.warnings:
            findings.append(Finding(
                severity="error", category="list_load_problem",
                item_type=lst.item_type, identifier="",
                message=w, source=lst.source_file,
            ))

        seen_ids: dict[str, int] = {}
        for item in lst.items:
            # Duplicate identifier within the same list.
            if item.identifier in seen_ids:
                findings.append(Finding(
                    severity="warning", category="duplicate_in_list",
                    item_type=lst.item_type, identifier=item.raw_identifier,
                    message=(f"Duplicate entry; first seen on row "
                             f"{seen_ids[item.identifier]}."),
                    source=lst.source_file,
                ))
            else:
                seen_ids[item.identifier] = item.row_number

            drawings = index.drawings_for(lst.item_type, item.identifier)
            matched_on_drawings.setdefault(lst.item_type, set())
            matched_on_drawings[lst.item_type] |= drawings

            if not drawings:
                findings.append(Finding(
                    severity="error", category="in_list_not_on_pid",
                    item_type=lst.item_type, identifier=item.raw_identifier,
                    message=("Listed but not found on any P&ID drawing."),
                    source=lst.source_file,
                ))
                continue

            # Item is declared against a specific drawing — verify it's there.
            if config.CHECK_DECLARED_PAGE and item.declared_drawing:
                if item.declared_drawing not in drawings:
                    findings.append(Finding(
                        severity="error", category="wrong_declared_drawing",
                        item_type=lst.item_type, identifier=item.raw_identifier,
                        message=(f"List says this is on {item.declared_drawing}, "
                                 f"but it was found on: {', '.join(sorted(drawings))}."),
                        source=lst.source_file,
                        drawing=item.declared_drawing,
                    ))

            # Appears on more than one drawing — often legitimate, flag as info.
            if len(drawings) > 1:
                findings.append(Finding(
                    severity="info", category="on_multiple_drawings",
                    item_type=lst.item_type, identifier=item.raw_identifier,
                    message=(f"Found on multiple drawings: "
                             f"{', '.join(sorted(drawings))}."),
                    source=lst.source_file,
                    drawing=", ".join(sorted(drawings)),
                ))

    # Reverse direction: tags on drawings that no list accounts for.
    list_ids_by_type: dict[str, set] = {}
    for lst in lists:
        list_ids_by_type.setdefault(lst.item_type, set())
        list_ids_by_type[lst.item_type] |= {it.identifier for it in lst.items}

    for item_type, tag_map in index.by_type.items():
        # Only flag the reverse when a list of that type was actually provided.
        if item_type not in list_ids_by_type:
            continue
        for tag, drawings in tag_map.items():
            if tag not in list_ids_by_type[item_type]:
                findings.append(Finding(
                    severity="error", category="on_pid_not_in_list",
                    item_type=item_type, identifier=tag,
                    message=("Appears on a P&ID but is missing from the "
                             f"{item_type} list."),
                    source="P&ID",
                    drawing=", ".join(sorted(drawings)),
                ))

    return findings


def summarize(findings: list[Finding]) -> dict:
    summary = {"error": 0, "warning": 0, "info": 0, "by_category": {}}
    for f in findings:
        summary[f.severity] = summary.get(f.severity, 0) + 1
        summary["by_category"][f.category] = summary["by_category"].get(f.category, 0) + 1
    return summary
