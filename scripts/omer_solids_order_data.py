# -*- coding: utf-8 -*-
"""Shared quantities and copy for Omer solid-interlock garment order."""
from __future__ import annotations

DATE = "11 September 2026"
OUT_PDF_NAME = "Omer_solid_interlock_garment_order.pdf"
OUT_XLSX_NAME = "Omer_solid_interlock_garment_order.xlsx"

COLORS = [
    {"key": "navy", "name": "BLUE NAVY", "code": "230416", "note": "Solid only"},
    {"key": "pink", "name": "PINK", "code": "216902", "note": "Solid only"},
    {"key": "gray", "name": "GRAY", "code": "131994", "note": "Light gray / solid gray — same color"},
    {"key": "offwhite", "name": "OFF WHITE", "code": "—", "note": "No code. Match previous production. Do not substitute."},
]

FABRIC = [
    ("Type", "Interlock"),
    ("Composition", "100% cotton"),
    ("Yarn", "30/1"),
    ("GSM", "220"),
    ("Brush", "One side"),
    ("Finish", "Solid color only"),
]

# A. Overall — new cut, never sewn. 0-3 / 3-6 / 6-12
OVERALL_QTY = {"0-3": 200, "3-6": 150, "6-12": 150}

# B. Pajama set — 202 already sewn + new shirt cut. 12-18 / 18-24 / 24-30
PAJAMA_QTY = 80
PAJAMA_SIZES = ("12-18", "18-24", "24-30")

# C. Baby set — 201 wrap + footed pants, both already sewn. 3-6 / 6-12
BABY_QTY = 80
BABY_SIZES = ("3-6", "6-12")

# D. Newborn set — same 201 + footed pants. 0-3 only
NEWBORN_QTY = {"navy": 120, "pink": 120, "gray": 160, "offwhite": 120}


def totals() -> dict[str, int]:
    overall = sum(OVERALL_QTY[s] * len(COLORS) for s in OVERALL_QTY)
    pajama = PAJAMA_QTY * len(PAJAMA_SIZES) * len(COLORS)
    baby = BABY_QTY * len(BABY_SIZES) * len(COLORS)
    newborn = sum(NEWBORN_QTY.values())
    return {
        "overall": overall,
        "pajama": pajama,
        "baby": baby,
        "newborn": newborn,
        "wrap_sets": baby + newborn,
        "units": overall + pajama + baby + newborn,
    }


PRODUCTS = [
    {
        "code": "A",
        "name": "Overall / footed romper",
        "short": "Overall",
        "construction": "Closed feet, double zipper",
        "sizes": "0-3 / 3-6 / 6-12",
        "unit": "pcs",
        "pattern": "NEW pattern — you have not sewn this cut yet",
        "models": "New overall / sleepsuit",
    },
    {
        "code": "B",
        "name": "Pajama set",
        "short": "Pajama set",
        "construction": "Long-sleeve shirt + open gatkes (202)",
        "sizes": "12-18 / 18-24 / 24-30",
        "unit": "sets",
        "pattern": "202 open pants — already sewn. Shirt — NEW pattern, not sewn yet",
        "models": "New shirt + 202",
    },
    {
        "code": "C",
        "name": "Baby set",
        "short": "Baby set",
        "construction": "201 envelope wrap bodysuit + footed pants",
        "sizes": "3-6 / 6-12",
        "unit": "sets",
        "pattern": "201 wrap + footed pants — both patterns already sewn",
        "models": "201 + footed pants",
    },
    {
        "code": "D",
        "name": "Newborn set",
        "short": "Newborn set",
        "construction": "201 envelope wrap bodysuit + footed pants",
        "sizes": "0-3 only",
        "unit": "sets",
        "pattern": "201 wrap + footed pants — both patterns already sewn",
        "models": "201 + footed pants",
    },
]
