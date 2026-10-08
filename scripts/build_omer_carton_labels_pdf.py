# -*- coding: utf-8 -*-
"""One-off carton labels PDF for Omer packing list 200-201-202.

Two A4 files (Baby Basic / Arye), same 3x5 label-sheet geometry as the app,
two identical labels per carton.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from optitex_analyzer.core.label_sheet import (  # noqa: E402
    BORDER_INSET,
    COLS,
    CONTENT_PAD,
    H_GAP,
    LABEL_H,
    LABEL_W,
    LOGO_BABY_BASIC_PATH,
    LOGO_PATH,
    PAGE_H,
    PAGE_W,
    PER_PAGE,
    ROWS,
    ROW_SHIFT_DOWN_MM,
    V_GAP,
    _FONT_BOLD,
    _FONT_LIGHT,
    _FONT_NAME,
    _draw_image_fit,
    _get_logo,
    _register_fonts,
    _shape,
)
from build_omer_shipment_payment import SUPPLIER_DIR, load_cartons  # noqa: E402

PACKING_NAME = "200-201-202.xlsx"
MODEL_NAMES = {
    200: "רגליות",
    201: "בגד גוף מעטפת",
    202: "גטקסים",
}
COPIES_PER_CARTON = 2
LOGO_H = 1.0 * cm
OUTPUTS = {
    "BABY BASIC": {
        "dest": SUPPLIER_DIR / "מדבקות_קרטונים_בייבי_בייסיק.pdf",
        "logo": LOGO_BABY_BASIC_PATH,
    },
    "ARYE": {
        "dest": SUPPLIER_DIR / "מדבקות_קרטונים_אריה.pdf",
        "logo": LOGO_PATH,
        "emphasize": True,
    },
}


def _find_packing_list() -> Path:
    for name in os.listdir(SUPPLIER_DIR):
        if name.startswith("~$"):
            continue
        if name.lower() == PACKING_NAME:
            return SUPPLIER_DIR / name
    raise FileNotFoundError(f"{PACKING_NAME} not found in {SUPPLIER_DIR}")


def _fit_size(c: canvas.Canvas, text: str, font: str, max_size: float, min_size: float, max_width: float) -> float:
    size = max_size
    while size > min_size and c.stringWidth(text, font, size) > max_width:
        size -= 0.4
    return size


def _draw_qty_line(c, mid_x, y, font, size, pieces, packets, max_w) -> None:
    """יחידות מימין ומארזים משמאל, בלי לערבב את סדר העברית והמספרים."""
    units = _shape(f"{pieces} יחידות")
    packs = _shape(f"{packets} מארזים")
    sep = "·"
    c.setFont(font, size)
    w_units = c.stringWidth(units, font, size)
    w_packs = c.stringWidth(packs, font, size)
    w_sep = c.stringWidth(sep, font, size)
    pad = 5
    total = w_units + w_packs + w_sep + pad * 2
    while size > 7.5 and total > max_w:
        size -= 0.3
        c.setFont(font, size)
        w_units = c.stringWidth(units, font, size)
        w_packs = c.stringWidth(packs, font, size)
        w_sep = c.stringWidth(sep, font, size)
        total = w_units + w_packs + w_sep + pad * 2
    right = mid_x + total / 2.0
    c.drawRightString(right, y, units)
    c.drawCentredString(right - w_units - pad - w_sep / 2.0, y, sep)
    c.drawRightString(right - w_units - pad * 2 - w_sep, y, packs)


def _draw_carton_label(c: canvas.Canvas, x_left: float, y_bottom: float, item: dict, logo, emphasize: bool = False) -> None:
    x0 = x_left + BORDER_INSET + CONTENT_PAD
    x1 = x_left + LABEL_W - BORDER_INSET - CONTENT_PAD
    y0 = y_bottom + BORDER_INSET + CONTENT_PAD
    y1 = y_bottom + LABEL_H - BORDER_INSET - CONTENT_PAD
    mid_x = (x0 + x1) / 2.0
    content_w = x1 - x0

    if logo is not None:
        _draw_image_fit(c, logo, x0, y1 - LOGO_H, content_w, LOGO_H, anchor="n")
        y1 -= LOGO_H + 0.06 * cm

    carton_no = int(item["carton"])
    model_name = MODEL_NAMES.get(int(item["model"]), str(item["model"]))
    size = str(item.get("size") or "").strip()
    size_text = size if emphasize else (f"{size} חודשים" if size else "")
    pieces = int(item.get("pieces") or 0)
    packets = int(item.get("packets") or 0)

    number_size = 22
    model_max = 22 if emphasize else 16
    size_max = 20 if emphasize else 14
    model_size = _fit_size(c, _shape(model_name), _FONT_BOLD, model_max, 12, content_w)
    size_size = _fit_size(c, _shape(size_text), _FONT_BOLD, size_max, 12, content_w)
    qty_size = 10
    gap = 2.2
    block_h = number_size + model_size + size_size + qty_size + gap * 3
    y = y0 + (y1 - y0 + block_h) / 2.0

    y -= number_size
    c.setFont(_FONT_BOLD, number_size)
    c.drawCentredString(mid_x, y, str(carton_no))
    y -= gap + model_size
    c.setFont(_FONT_BOLD, model_size)
    c.drawCentredString(mid_x, y, _shape(model_name))
    y -= gap + size_size
    c.setFont(_FONT_BOLD, size_size)
    c.drawCentredString(mid_x, y, _shape(size_text))
    y -= gap + qty_size
    _draw_qty_line(c, mid_x, y, _FONT_LIGHT, qty_size, pieces, packets, content_w)


def _cell_origin(index: int) -> tuple[float, float]:
    pos = index % PER_PAGE
    row = pos // COLS
    col = pos % COLS
    left_margin = (PAGE_W - (COLS * LABEL_W + (COLS - 1) * H_GAP)) / 2.0
    top_margin = (PAGE_H - (ROWS * LABEL_H + (ROWS - 1) * V_GAP)) / 2.0
    x_left = left_margin + col * (LABEL_W + H_GAP)
    y_bottom = PAGE_H - top_margin - row * (LABEL_H + V_GAP) - LABEL_H
    shift_mm = ROW_SHIFT_DOWN_MM[row] if row < len(ROW_SHIFT_DOWN_MM) else 0.0
    y_bottom -= shift_mm * mm
    return x_left, y_bottom


def build_brand_pdf(cartons: list[dict], dest: Path, logo_path: str, emphasize: bool = False) -> int:
    items: list[dict] = []
    for carton in sorted(cartons, key=lambda c: int(c["carton"])):
        items.extend([carton] * COPIES_PER_CARTON)
    if not items:
        raise ValueError(f"אין קרטונים ל-{dest.name}")

    _register_fonts()
    logo = _get_logo(logo_path)
    if logo is None:
        raise FileNotFoundError(f"לוגו לא נמצא: {logo_path}")
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.stem + ".__tmp.pdf")
    c = canvas.Canvas(str(tmp), pagesize=(PAGE_W, PAGE_H))
    for idx, item in enumerate(items):
        if idx > 0 and idx % PER_PAGE == 0:
            c.showPage()
        x_left, y_bottom = _cell_origin(idx)
        _draw_carton_label(c, x_left, y_bottom, item, logo, emphasize=emphasize)
    c.showPage()
    c.save()
    try:
        os.replace(tmp, dest)
    except PermissionError:
        alt = dest.with_name(dest.stem + "_מעודכן.pdf")
        os.replace(tmp, alt)
        print(f"הקובץ פתוח, נשמר במקום: {alt.name}")
        return len(items)
    return len(items)


def main() -> None:
    packing = _find_packing_list()
    cartons = load_cartons(packing)
    by_brand: dict[str, list[dict]] = {key: [] for key in OUTPUTS}
    unknown = []
    for item in cartons:
        brand = item.get("brand") or ""
        if brand in by_brand:
            by_brand[brand].append(item)
        else:
            unknown.append(item)
    if unknown:
        brands = sorted({str(i.get("brand")) for i in unknown})
        raise RuntimeError(f"מותגים לא מזוהים: {brands}")

    for brand, spec in OUTPUTS.items():
        dest = spec["dest"]
        count = build_brand_pdf(by_brand[brand], dest, spec["logo"], emphasize=bool(spec.get("emphasize")))
        pages = (count + PER_PAGE - 1) // PER_PAGE
        print(f"{dest.name}: cartons={len(by_brand[brand])} labels={count} pages={pages}")


if __name__ == "__main__":
    main()
