# -*- coding: utf-8 -*-
"""One-off: add Akdem / Bilal as supplier 5 and attach socks-order documents."""
from __future__ import annotations

import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)

from optitex_analyzer.core.data_processor import DataProcessor  # noqa: E402

BILAL_DIR = Path(r"C:\Users\levyn\OneDrive\שולחן העבודה\ביללאל")
NOTES = (
    "בדים אינטרלוק/ריב. גרביים דרך Stilteks (בורסה): הזמנה אחרונה 08/2022 — "
    "1,000 מארזי חמישיות לבן @2.20$, נשלחו 996 מארזים (4,980 זוגות) 22.02.2023"
)

NAVY = PatternFill("solid", fgColor="1B2A4A")
ALT = PatternFill("solid", fgColor="F3EFE8")
PAPER = PatternFill("solid", fgColor="FAF7F2")
SECTION = PatternFill("solid", fgColor="E8EEF6")
WHITE = PatternFill("solid", fgColor="FFFFFF")
THIN = Border(
    left=Side(style="thin", color="D6D3D1"),
    right=Side(style="thin", color="D6D3D1"),
    top=Side(style="thin", color="D6D3D1"),
    bottom=Side(style="thin", color="D6D3D1"),
)
TITLE = Font(name="Calibri", bold=True, size=14, color="1B2A4A")
HEAD = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
BODY = Font(name="Calibri", size=10, color="1C1917")
BOLD = Font(name="Calibri", bold=True, size=10, color="1C1917")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _find_under(folder: Path, filename: str) -> Path:
    direct = folder / filename
    if direct.is_file():
        return direct
    matches = list(folder.rglob(filename))
    if matches:
        return matches[0]
    raise FileNotFoundError(f"Missing {filename} under {folder}")


def _copy_bytes(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(src.read_bytes())


def _paint(cell, *, fill=None, font=None, align=None) -> None:
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font
    if align is not None:
        cell.alignment = align
    cell.border = THIN


def build_history_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Socks order history"
    ws.sheet_view.showGridLines = False
    widths = [14, 28, 16, 12, 16, 16, 55]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:G1")
    c = ws.cell(row=1, column=1, value="BILAL / STILTEKS — last baby socks order")
    c.font = TITLE
    c.fill = PAPER
    ws.row_dimensions[1].height = 24

    headers = ["Date", "Document", "Packs", "Pairs", "USD / pack", "Amount USD", "Note"]
    for col, title in enumerate(headers, 1):
        _paint(ws.cell(row=3, column=col, value=title), fill=NAVY, font=HEAD, align=CENTER)
    ws.row_dimensions[3].height = 22

    rows = [
        (datetime(2022, 8, 16), "Stilteks proforma", 1000, 5000, 2.20, 2200.00, "5'li BEBE ÇORAP, white. Ordered via Bilal."),
        (datetime(2022, 8, 31), "SWIFT payment", None, None, None, 2200.00, "Paid in full to Stilteks Ltd, QNB Finansbank."),
        (datetime(2023, 2, 22), "Stilteks commercial invoice #3", 996, 4980, 2.20, 2191.20, "White socks. STL2023000000029."),
        (datetime(2023, 2, 22), "Packing list", 996, 4980, None, None, "5 cartons (koli). Net 54.75 kg, gross 58.75 kg."),
    ]
    for i, (dt, doc, packs, pairs, unit, amount, note) in enumerate(rows):
        r = 4 + i
        fill = ALT if i % 2 else WHITE
        values = [dt, doc, packs, pairs, unit, amount, note]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=r, column=col, value=val)
            _paint(cell, fill=fill, font=BODY, align=CENTER if col != 7 else LEFT)
            if col == 1:
                cell.number_format = "DD.MM.YYYY"
            if col in (3, 4) and val is not None:
                cell.number_format = "#,##0"
            if col in (5, 6) and val is not None:
                cell.number_format = '"$"#,##0.00'
        ws.row_dimensions[r].height = 20

    ws.merge_cells("A9:G9")
    h = ws.cell(row=9, column=1, value="Stilteks (socks factory, Bursa) — not Akdem")
    h.font = BOLD
    h.fill = SECTION
    for col in range(2, 8):
        ws.cell(row=9, column=col).fill = SECTION
        ws.cell(row=9, column=col).border = THIN
    ws.row_dimensions[9].height = 22

    details = [
        ("Company", "STILTEKS OTOMOTIV GIDA INSAAT CORAP TEKSTIL TURIZM SAN. VE TIC. LTD. STI."),
        ("Address", "Kazimkarabekir Mh. 1. Ula Sk. No:4/2, Yildirim / Bursa, Turkey"),
        ("Phone / fax", "+90 224 360 18 38"),
        ("Bank", "QNB Finansbank, Yavuz Selim branch"),
        ("IBAN (USD)", "TR36 0011 1000 0000 0027 5565 90"),
        ("SWIFT", "FNNBTRISXXX"),
        ("Buyer", "ARYE MOSHE LEVY, 6 Hakishon St, Tel Aviv, Israel, ZIP 6609306, VAT 52064219"),
        ("Channel", "Ordered through Bilal Gursoy (Akdem, Kestel/Bursa). No later socks order in the WhatsApp chat."),
    ]
    for i, (label, value) in enumerate(details):
        r = 10 + i
        _paint(ws.cell(row=r, column=1, value=label), fill=ALT, font=BOLD, align=LEFT)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        _paint(ws.cell(row=r, column=2, value=value), fill=WHITE, font=BODY, align=LEFT)
        for col in range(3, 8):
            ws.cell(row=r, column=col).fill = WHITE
            ws.cell(row=r, column=col).border = THIN
        ws.row_dimensions[r].height = 18

    ws.print_area = "A1:G17"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def attach(dp: DataProcessor, supplier_id: int, src: Path, dest_name: str, original_name: str) -> None:
    tmp_dir = ROOT / "exports" / "_bilal_docs_tmp"
    tmp_path = tmp_dir / dest_name
    _copy_bytes(src, tmp_path)
    rec = dp.add_supplier_document(supplier_id, str(tmp_path))
    rec["original_name"] = original_name
    rec["filename"] = dest_name
    dest = Path(dp.supplier_document_path(supplier_id, dest_name))
    if dest.resolve() != tmp_path.resolve():
        # add_supplier_document may have kept dest_name; if it timestamped, rename back
        copied = Path(dp.supplier_document_path(supplier_id, rec.get("filename") or dest_name))
        if copied != dest and copied.is_file():
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                dest.unlink()
            shutil.move(str(copied), str(dest))
        rec["filename"] = dest_name
    dp.save_suppliers()


def main() -> None:
    dp = DataProcessor()
    existing = next(
        (s for s in dp.suppliers if (s.get("business_name") or "").lower().startswith("akdem")),
        None,
    )
    if existing:
        supplier_id = int(existing["id"])
        existing["notes"] = NOTES
        existing["first_name"] = existing.get("first_name") or "בילאל"
        existing["address"] = existing.get("address") or "Kestel, Bursa, Turkey"
        dp.save_suppliers()
        print(f"supplier already exists id={supplier_id}")
    else:
        supplier_id = dp.add_supplier(
            "Akdem / Bilal Gursoy",
            "",
            "Kestel, Bursa, Turkey",
            "",
            NOTES,
            "בילאל",
        )
        print(f"added supplier id={supplier_id}")

    history_path = ROOT / "exports" / "Bilal_socks_order_history.xlsx"
    build_history_xlsx(history_path)

    chats = [p for p in BILAL_DIR.iterdir() if p.is_file() and "WhatsApp" in p.name and "Bilal" in p.name]
    if not chats:
        raise FileNotFoundError("WhatsApp chat with Bilal not found")
    chat = chats[0]
    files = [
        (
            chat,
            "WhatsApp_chat_Bilal_Akadem.txt",
            chat.name,
        ),
        (
            _find_under(BILAL_DIR / "02_חשבוניות_ופרופורמות", "FATURA-1.pdf"),
            "Stilteks_proforma_16.08.2022.pdf",
            "FATURA-1.pdf",
        ),
        (
            _find_under(BILAL_DIR / "02_חשבוניות_ופרופורמות", "ınvoice.pdf"),
            "Stilteks_proforma_22.02.2023.pdf",
            "ınvoice.pdf",
        ),
        (
            _find_under(BILAL_DIR / "07_אחר", "ARYE MOSHE LEVY İNVOİCE.pdf"),
            "Stilteks_commercial_invoice_22.02.2023.pdf",
            "ARYE MOSHE LEVY İNVOİCE.pdf",
        ),
        (
            _find_under(BILAL_DIR / "03_רשימות_משקל_Packing", "çeki listesi (2).pdf"),
            "Stilteks_packing_list_22.02.2023.pdf",
            "çeki listesi (2).pdf",
        ),
        (
            _find_under(BILAL_DIR / "04_SWIFT_ובנק", "CCF31082022_0001.pdf"),
            "Stilteks_SWIFT_31.08.2022.pdf",
            "CCF31082022_0001.pdf",
        ),
        (
            history_path,
            "Bilal_socks_order_history.xlsx",
            "Bilal_socks_order_history.xlsx",
        ),
    ]

    supplier = dp.get_supplier(supplier_id)
    already = {(d.get("filename") or "") for d in (supplier.get("documents") or [])}
    for src, dest_name, original in files:
        if dest_name in already:
            print(f"skip existing {dest_name}")
            continue
        attach(dp, supplier_id, src, dest_name, original)
        print(f"attached {dest_name}")

    tmp_dir = ROOT / "exports" / "_bilal_docs_tmp"
    if tmp_dir.is_dir():
        shutil.rmtree(tmp_dir, ignore_errors=True)

    # Keep a copy of the history summary in exports for convenience
    print(history_path)
    dest_dir = Path(dp.supplier_documents_dir(supplier_id))
    print(dest_dir)
    for p in sorted(dest_dir.iterdir()):
        print(" ", p.name, p.stat().st_size)


if __name__ == "__main__":
    main()
