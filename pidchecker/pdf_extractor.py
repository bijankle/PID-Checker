"""
Extract, per P&ID page:
  - the drawing / P&ID number (from the title block when possible), and
  - every tag (line / valve / tie-in / equipment) appearing on that page.

Native/vector PDFs are read directly with PyMuPDF (fast, exact). Pages with no
extractable text are treated as scanned and fall back to OCR via pytesseract,
which is optional: if it (or the tesseract binary) is unavailable we record the
page as un-OCR'd rather than crashing.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import fitz  # PyMuPDF

import config


@dataclass
class PageResult:
    page_index: int                       # 0-based page index in the PDF
    drawing_number: Optional[str]          # extracted P&ID number, or None
    drawing_number_source: str             # "title_block" | "page_scan" | "none"
    used_ocr: bool                         # whether OCR was needed for this page
    tags: dict[str, set] = field(default_factory=dict)  # item_type -> {tags}
    raw_text: str = ""                     # full page text (for debugging)


def _compile_patterns() -> dict:
    return {name: re.compile(pat) for name, pat in config.TAG_PATTERNS.items()}


def _find_tags(text: str, patterns: dict) -> dict[str, set]:
    found: dict[str, set] = {}
    upper = text.upper()
    for name, rx in patterns.items():
        matches = {config.normalize_tag(m.group(0)) for m in rx.finditer(upper)}
        found[name] = {m for m in matches if m}
    return found


def _title_block_text(page: "fitz.Page") -> str:
    """Return text contained in the configured title-block region of the page."""
    rect = page.rect
    tb = config.TITLE_BLOCK
    clip = fitz.Rect(
        rect.x0 + tb["x0_frac"] * rect.width,
        rect.y0 + tb["y0_frac"] * rect.height,
        rect.x0 + tb["x1_frac"] * rect.width,
        rect.y0 + tb["y1_frac"] * rect.height,
    )
    return page.get_text("text", clip=clip)


def _extract_drawing_number(title_text: str) -> Optional[str]:
    """Pick the most likely drawing number from title-block text.

    Strategy: if a known label ("DWG NO" etc.) is present, prefer the first
    number-looking token after it; otherwise return the first token matching the
    drawing-number pattern.
    """
    rx = re.compile(config.DRAWING_NUMBER_PATTERN)
    upper = title_text.upper()

    for label in config.DRAWING_NUMBER_LABELS:
        idx = upper.find(label.upper())
        if idx != -1:
            after = upper[idx + len(label):]
            m = rx.search(after)
            if m:
                return config.normalize_tag(m.group(0))

    m = rx.search(upper)
    return config.normalize_tag(m.group(0)) if m else None


def _ocr_page(page: "fitz.Page") -> str:
    """Best-effort OCR of a rasterized page. Returns '' if OCR is unavailable."""
    try:
        import pytesseract  # noqa: WPS433 (local import: optional dependency)
        from PIL import Image
        import io

        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return pytesseract.image_to_string(img)
    except Exception:
        # tesseract binary or pytesseract/Pillow missing, or OCR failed.
        return ""


def extract_pdf(path: str, ocr_text_threshold: int = 20) -> list[PageResult]:
    """Process every page of a P&ID PDF.

    ocr_text_threshold: if a page yields fewer than this many characters of
    native text it is considered scanned and we attempt OCR.
    """
    patterns = _compile_patterns()
    results: list[PageResult] = []

    with fitz.open(path) as doc:
        for i, page in enumerate(doc):
            native_text = page.get_text("text")
            used_ocr = False

            if len(native_text.strip()) < ocr_text_threshold:
                ocr_text = _ocr_page(page)
                if ocr_text.strip():
                    native_text = ocr_text
                    used_ocr = True

            # Drawing number: try title block first, then whole page.
            tb_text = _title_block_text(page) if not used_ocr else native_text
            dwg = _extract_drawing_number(tb_text)
            source = "title_block"
            if not dwg:
                dwg = _extract_drawing_number(native_text)
                source = "page_scan" if dwg else "none"

            results.append(
                PageResult(
                    page_index=i,
                    drawing_number=dwg,
                    drawing_number_source=source,
                    used_ocr=used_ocr,
                    tags=_find_tags(native_text, patterns),
                    raw_text=native_text,
                )
            )

    return results
