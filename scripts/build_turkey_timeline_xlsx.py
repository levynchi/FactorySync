# -*- coding: utf-8 -*-
"""Build Turkey shipments/payments timeline Excel from curated events."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from turkey_timeline_data import (
    AKDEM_LEDGER_SUMMARY,
    EVENTS,
    GAPS,
    OMER_FIRST_SHIPMENT,
    SOURCES,
    invoices,
    ils_shipping,
    shipments,
    usd_payment_rows,
    usd_payments,
)

ROOT = Path(__file__).resolve().parents[1]
OUT_PATH = ROOT / "exports" / "turkey_shipments_payments_timeline.xlsx"
SHIPPING_COSTS = ROOT / "shipping_costs.json"
AKDEM_LEDGER = Path(
    r"C:\Users\levyn\OneDrive\שולחן העבודה\ביללאל\00_כרטסות_ובאלנסים\ARYE MOSHE balance.xlsx"
)

NAVY = PatternFill("solid", fgColor="1B2A4A")
ALT = PatternFill("solid", fgColor="F8FAFC")
PAPER = PatternFill("solid", fgColor="FFFFFF")
SECTION = PatternFill("solid", fgColor="E8EEF6")
TOTAL = PatternFill("solid", fgColor="DBEAFE")
WARN = PatternFill("solid", fgColor="FEF3C7")
GOOD = PatternFill("solid", fgColor="D1FAE5")
BAD = PatternFill("solid", fgColor="FEE2E2")
KIND_FILL = {
    "תשלום ספק": PatternFill("solid", fgColor="DCFCE7"),
    "תשלום סחורה דרך משלח": PatternFill("solid", fgColor="BBF7D0"),
    "תשלום שילוח/עמילות": PatternFill("solid", fgColor="DBEAFE"),
    "משלוח יצא": PatternFill("solid", fgColor="FDE68A"),
    "משלוח הגיע": PatternFill("solid", fgColor="FCD34D"),
    "חשבונית": PatternFill("solid", fgColor="E9D5FF"),
    "פרופורמה": PatternFill("solid", fgColor="F3E8FF"),
    "מסמך": PatternFill("solid", fgColor="E2E8F0"),
}
CERT_FILL = {
    "ודאי": GOOD,
    "סביר": WARN,
    "חלקי": BAD,
}
THIN = Border(
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
    bottom=Side(style="thin", color="E2E8F0"),
)
TITLE = Font(name="Calibri", bold=True, size=16, color="1B2A4A")
SUB = Font(name="Calibri", size=11, color="475569")
HEAD = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
BODY = Font(name="Calibri", size=10, color="1C1917")
BOLD = Font(name="Calibri", bold=True, size=10, color="1C1917")
NOTE = Font(name="Calibri", italic=True, size=9, color="64748B")
WHITE_BOLD = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="right", vertical="center", wrap_text=True, readingOrder=2)
LEFT_LTR = Alignment(horizontal="left", vertical="center", wrap_text=True)
NUM = Alignment(horizontal="center", vertical="center")

USD_FMT = '#,##0.00'
ILS_FMT = '#,##0.00'
DATE_FMT = "YYYY-MM-DD"


def _paint(cell, *, fill=None, font=None, align=None, border=True) -> None:
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font
    if align is not None:
        cell.alignment = align
    if border:
        cell.border = THIN


def _rtl(ws: Worksheet) -> None:
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.horizontalCentered = True
    ws.oddFooter.right.text = "עמוד &P מתוך &N"
    ws.oddFooter.left.text = "סודי — אריה בגדי תינוקות / משה לוי"


def _widths(ws: Worksheet, widths: list[float]) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _banner(ws: Worksheet, ncols: int, title: str, subtitle: str) -> int:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1, title)
    _paint(c, fill=NAVY, font=Font(name="Calibri", bold=True, size=16, color="FFFFFF"), align=LEFT, border=False)
    for col in range(2, ncols + 1):
        _paint(ws.cell(1, col), fill=NAVY, border=False)
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(2, 1, subtitle)
    _paint(c, fill=SECTION, font=SUB, align=LEFT, border=False)
    for col in range(2, ncols + 1):
        _paint(ws.cell(2, col), fill=SECTION, border=False)
    ws.row_dimensions[2].height = 22
    return 4


def _headers(ws: Worksheet, row: int, headers: list[str]) -> None:
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row, col, h)
        _paint(cell, fill=NAVY, font=HEAD, align=CENTER)
    ws.row_dimensions[row].height = 22
    ws.freeze_panes = f"A{row + 1}"
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(headers))}{row}"


def _sort_key(e: dict) -> tuple:
    return (e.get("date") or "9999", e.get("kind") or "")


def _write_event_row(ws: Worksheet, row: int, e: dict, cols: list[str]) -> None:
    mapping = {
        "תאריך": e.get("date"),
        "סוג": e.get("kind"),
        "גורם": e.get("party"),
        "תיאור": e.get("description"),
        "סכום": e.get("amount"),
        "מטבע": e.get("currency") or None,
        "₪": e.get("ils"),
        "כמות": e.get("qty") or None,
        "מקור": e.get("source"),
        "ודאות": e.get("certainty"),
        "הערות": e.get("notes") or None,
    }
    kind_fill = KIND_FILL.get(e.get("kind") or "", PAPER)
    fill = ALT if row % 2 == 0 else kind_fill
    for col, name in enumerate(cols, 1):
        val = mapping.get(name)
        cell = ws.cell(row, col, val if val not in ("", None) else None)
        align = LEFT if name in ("תיאור", "מקור", "הערות", "גורם", "כמות") else CENTER
        _paint(cell, fill=fill, font=BODY, align=align)
        if name == "סכום" and isinstance(val, (int, float)):
            cell.number_format = USD_FMT if (e.get("currency") or "") == "USD" else ILS_FMT
        if name == "₪" and isinstance(val, (int, float)):
            cell.number_format = ILS_FMT
        if name == "ודאות":
            _paint(cell, fill=CERT_FILL.get(e.get("certainty") or "", fill), font=BOLD, align=CENTER)
        if name == "סוג":
            _paint(cell, fill=kind_fill, font=BOLD, align=CENTER)
    ws.row_dimensions[row].height = 36


def _extend_filter(ws: Worksheet, header_row: int, last_row: int, ncols: int) -> None:
    if last_row <= header_row:
        return
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ncols)}{last_row}"


def sheet_timeline(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "ציר זמן"
    _rtl(ws)
    cols = ["תאריך", "סוג", "גורם", "תיאור", "סכום", "מטבע", "₪", "כמות", "ודאות", "מקור", "הערות"]
    _widths(ws, [12, 22, 22, 62, 14, 8, 12, 28, 10, 55, 40])
    built = datetime.now().strftime("%d.%m.%Y %H:%M")
    _banner(
        ws,
        len(cols),
        "משלוחים ותשלומים מול טורקיה — ציר זמן",
        f"אריה בגדי תינוקות / משה לוי  ·  נבנה {built}  ·  {len(EVENTS)} אירועים מאוחדים  ·  מקורות: מייל arye.baby, וואטסאפ, כרטסת Akdem, תוכנה",
    )
    header_row = 4
    _headers(ws, header_row, cols)
    rows = sorted(EVENTS, key=_sort_key)
    r = header_row + 1
    for e in rows:
        _write_event_row(ws, r, e, cols)
        r += 1
    _extend_filter(ws, header_row, r - 1, len(cols))
    note_row = r + 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=len(cols))
    _paint(
        ws.cell(
            note_row,
            1,
            "צבע לפי סוג אירוע. ודאות: ירוק=SWIFT/כרטסת/מייל בנק, צהוב=הוראת העברה בלי SWIFT, אדום=אזכור בצ'אט בלבד. "
            "יבוא סין/וייטנאם לא נכלל. אותה העברה שמופיעה במייל+וואטסאפ+כרטסת מוזגה לשורה אחת.",
        ),
        fill=SECTION,
        font=NOTE,
        align=LEFT,
        border=False,
    )


def _sum_amount(items: list[dict], currency: str) -> float:
    total = 0.0
    for e in items:
        if e.get("currency") == currency and e.get("amount"):
            total += float(e["amount"])
    return total


def sheet_usd(wb: Workbook) -> None:
    ws = wb.create_sheet("תשלומים לספקים USD")
    _rtl(ws)
    cols = ["תאריך", "סוג", "גורם", "תיאור", "סכום", "מטבע", "ודאות", "מקור", "הערות"]
    _widths(ws, [12, 24, 22, 70, 14, 8, 10, 55, 40])
    items = sorted(usd_payment_rows(), key=_sort_key)
    counted = usd_payments()
    total = _sum_amount(counted, "USD")
    _banner(
        ws,
        len(cols),
        "תשלומים לספקים ולמשלחים בטורקיה (USD)",
        f"{len(counted)} תשלומים שנספרים (+{len(items) - len(counted)} כפילות/הוחזרה בגיליון)  ·  סה״כ נטו: {total:,.2f} USD  ·  הוראות 2020–2023 לדיסקונט ללא SWIFT בגוף המייל מסומנות «סביר»",
    )
    header_row = 4
    _headers(ws, header_row, cols)
    r = header_row + 1
    for e in items:
        _write_event_row(ws, r, e, cols)
        r += 1
    _paint(ws.cell(r, 1, "סה״כ (ללא ההעברה שהוחזרה)"), fill=TOTAL, font=BOLD, align=LEFT)
    for col in range(2, 5):
        _paint(ws.cell(r, col), fill=TOTAL, border=True)
    cell = ws.cell(r, 5, total)
    cell.number_format = USD_FMT
    _paint(cell, fill=TOTAL, font=BOLD, align=CENTER)
    for col in range(6, len(cols) + 1):
        _paint(ws.cell(r, col), fill=TOTAL, border=True)
    _extend_filter(ws, header_row, r - 1, len(cols))


def sheet_ils(wb: Workbook) -> None:
    ws = wb.create_sheet("תשלומי שילוח ועמילות")
    _rtl(ws)
    cols = ["תאריך", "גורם", "תיאור", "סכום", "מטבע", "₪", "כמות", "ודאות", "מקור", "הערות"]
    _widths(ws, [12, 22, 70, 14, 8, 12, 28, 10, 55, 40])
    items = sorted(ils_shipping(), key=_sort_key)
    total_ils = sum(float(e["ils"] or e["amount"] or 0) for e in items if (e.get("ils") or e.get("amount")))
    _banner(
        ws,
        len(cols),
        "תשלומי שילוח, עמילות מכס והובלה פנים-ארצית על משלוחי טורקיה",
        f"{len(items)} שורות  ·  סה״כ מתועד ₪ {total_ils:,.2f}  ·  כולל חיים נתנאל, און אקספרס, חמודי, סאקיס, אקספרסליין",
    )
    header_row = 4
    _headers(ws, header_row, cols)
    r = header_row + 1
    for e in items:
        _write_event_row(ws, r, e, cols)
        r += 1
    _paint(ws.cell(r, 1, "סה״כ ₪ מתועד"), fill=TOTAL, font=BOLD, align=LEFT)
    for col in range(2, 6):
        _paint(ws.cell(r, col), fill=TOTAL)
    cell = ws.cell(r, 6, total_ils)
    cell.number_format = ILS_FMT
    _paint(cell, fill=TOTAL, font=BOLD, align=CENTER)
    for col in range(7, len(cols) + 1):
        _paint(ws.cell(r, col), fill=TOTAL)
    _extend_filter(ws, header_row, r - 1, len(cols))


def sheet_shipments(wb: Workbook) -> None:
    ws = wb.create_sheet("משלוחים")
    _rtl(ws)
    cols = ["תאריך", "סוג", "גורם", "תיאור", "כמות", "ודאות", "מקור", "הערות"]
    _widths(ws, [12, 16, 24, 75, 32, 10, 55, 40])
    items = sorted(shipments(), key=_sort_key)
    _banner(
        ws,
        len(cols),
        "משלוחים מטורקיה — יציאה והגעה/שחרור",
        f"{len(items)} אירועים  ·  ספקים: Akdem/IREMAK, Nibby, Stilteks  ·  משלחים: Mega Alfa, Nathaniel, On Express, Adel/חמודי, Sakis, Expressline",
    )
    header_row = 4
    _headers(ws, header_row, cols)
    r = header_row + 1
    for e in items:
        _write_event_row(ws, r, e, cols)
        r += 1
    _extend_filter(ws, header_row, r - 1, len(cols))


def sheet_invoices(wb: Workbook) -> None:
    ws = wb.create_sheet("חשבוניות ופרופורמות")
    _rtl(ws)
    cols = ["תאריך", "סוג", "גורם", "תיאור", "סכום", "מטבע", "כמות", "ודאות", "מקור", "הערות"]
    _widths(ws, [12, 12, 22, 70, 14, 8, 28, 10, 55, 40])
    items = sorted(invoices(), key=_sort_key)
    _banner(
        ws,
        len(cols),
        "חשבוניות ופרופורמות מטורקיה",
        "Akdem (ישיר / Adel / Sakis / Express Line), Nibby, Stilteks, IPTAS מגבת",
    )
    header_row = 4
    _headers(ws, header_row, cols)
    r = header_row + 1
    for e in items:
        _write_event_row(ws, r, e, cols)
        r += 1
    _extend_filter(ws, header_row, r - 1, len(cols))


def _load_akdem_ledger() -> list[tuple]:
    if not AKDEM_LEDGER.exists():
        return []
    wb = load_workbook(AKDEM_LEDGER, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if not row or row[0] is None:
            continue
        date = row[0]
        if hasattr(date, "strftime"):
            date_s = date.strftime("%Y-%m-%d")
        else:
            date_s = str(date)[:10]
        rows.append(
            (
                date_s,
                row[1],
                row[2],
                row[3],
                row[4],
                row[6],
                row[7],
            )
        )
    return rows


def sheet_balances(wb: Workbook) -> None:
    ws = wb.create_sheet("יתרות")
    _rtl(ws)
    _widths(ws, [42, 16, 16, 16, 55])
    _banner(
        ws,
        5,
        "יתרות פתוחות מול Akdem — טענת הספק מול עמדתנו",
        "מקור: ARYE MOSHE balance.xlsx (יולי 2026) + מכתב/וואטסאפ 15.07.2026",
    )
    headers = ["חברה / סעיף", "חויב USD", "שולם USD", "הפרש USD", "הערה"]
    _headers(ws, 4, headers)
    r = 5
    for name, charged, paid, diff in AKDEM_LEDGER_SUMMARY["our_split"]:
        vals = [name, charged, paid, diff, ""]
        fill = ALT if r % 2 == 0 else PAPER
        for col, v in enumerate(vals, 1):
            cell = ws.cell(r, col, v)
            _paint(cell, fill=fill, font=BODY, align=CENTER if col > 1 else LEFT)
            if col in (2, 3, 4) and isinstance(v, (int, float)):
                cell.number_format = USD_FMT
        r += 1
    s = AKDEM_LEDGER_SUMMARY
    for col, v in enumerate(
        ["סה״כ כרטסת Akdem", s["charges_usd"], s["paid_usd"], s["claimed_balance_usd"], "יתרה שהספק טוען"],
        1,
    ):
        cell = ws.cell(r, col, v)
        _paint(cell, fill=TOTAL, font=BOLD, align=CENTER if col > 1 else LEFT)
        if col in (2, 3, 4):
            cell.number_format = USD_FMT
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    _paint(ws.cell(r, 1, "עמדת אריה — מה באמת פתוח"), fill=NAVY, font=WHITE_BOLD, align=LEFT, border=False)
    for col in range(2, 6):
        _paint(ws.cell(r, col), fill=NAVY, border=False)
    r += 1
    _headers(ws, r, ["סעיף", "סכום פתוח USD", "", "", "פירוט"])
    r += 1
    for name, amt in s["our_view_open"]:
        _paint(ws.cell(r, 1, name), fill=WARN, font=BODY, align=LEFT)
        cell = ws.cell(r, 2, amt)
        cell.number_format = USD_FMT
        _paint(cell, fill=WARN, font=BOLD, align=CENTER)
        for col in range(3, 6):
            _paint(ws.cell(r, col), fill=WARN)
        r += 1
    our_open = sum(a for _, a in s["our_view_open"])
    _paint(ws.cell(r, 1, "סה״כ לפי עמדתנו"), fill=BAD, font=BOLD, align=LEFT)
    cell = ws.cell(r, 2, our_open)
    cell.number_format = USD_FMT
    _paint(cell, fill=BAD, font=BOLD, align=CENTER)
    for col in range(3, 6):
        _paint(ws.cell(r, col), fill=BAD)
    r += 2
    o = OMER_FIRST_SHIPMENT
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    _paint(
        ws.cell(r, 1, f"משלוח ראשון עומר {o['date']} — {o['pcs']:,} יח' / {o['cartons']} קרטונים / {o['amount_usd']:,.2f}$"),
        fill=NAVY,
        font=WHITE_BOLD,
        align=LEFT,
        border=False,
    )
    for col in range(2, 6):
        _paint(ws.cell(r, col), fill=NAVY, border=False)
    r += 1
    _paint(ws.cell(r, 1, "מקדמה שדוברה בוואטסאפ"), fill=PAPER, font=BODY, align=LEFT)
    cell = ws.cell(r, 2, o["advance_claimed_usd"])
    cell.number_format = USD_FMT
    _paint(cell, fill=PAPER, font=BODY, align=CENTER)
    r += 1
    _paint(ws.cell(r, 1, "SWIFT שתועד ל-Express Line"), fill=PAPER, font=BODY, align=LEFT)
    cell = ws.cell(r, 2, o["advance_swift_usd"])
    cell.number_format = USD_FMT
    _paint(cell, fill=PAPER, font=BODY, align=CENTER)
    r += 1
    _paint(ws.cell(r, 1, "פער מקדמה"), fill=WARN, font=BOLD, align=LEFT)
    cell = ws.cell(r, 2, o["advance_claimed_usd"] - o["advance_swift_usd"])
    cell.number_format = USD_FMT
    _paint(cell, fill=WARN, font=BOLD, align=CENTER)
    r += 2
    ledger = _load_akdem_ledger()
    if ledger:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        _paint(
            ws.cell(r, 1, "העתק כרטסת Akdem (ARYE MOSHE) — חיובים ותקבולים"),
            fill=NAVY,
            font=WHITE_BOLD,
            align=LEFT,
            border=False,
        )
        for col in range(2, 6):
            _paint(ws.cell(r, col), fill=NAVY, border=False)
        r += 1
        led_headers = ["תאריך", "סדרה / מספר", "סוג מסמך", "תיאור", "חיוב USD"]
        # put credit in col 3 wait - we have 5 cols. Use: date, type, desc, debit, credit
        # Actually 7 fields compacted: date, seri-sira, type, desc, debit, and we need credit.
        # Let's use extra columns 2-5: type, desc, debit, credit - and put seri in desc
        _headers(ws, r, ["תאריך", "סוג מסמך", "תיאור בכרטסת", "חיוב USD", "זכות USD"])
        led_header = r
        r += 1
        for date_s, seri, sira, typ, desc, debit, credit in ledger:
            label = f"{seri or ''} {sira or ''}".strip()
            desc_s = f"{desc or ''} ({label})".strip()
            fill = BAD if debit else GOOD
            _paint(ws.cell(r, 1, date_s), fill=fill, font=BODY, align=CENTER)
            _paint(ws.cell(r, 2, typ), fill=fill, font=BODY, align=CENTER)
            _paint(ws.cell(r, 3, desc_s), fill=fill, font=BODY, align=LEFT)
            c4 = ws.cell(r, 4, debit if debit else None)
            c4.number_format = USD_FMT
            _paint(c4, fill=fill, font=BODY, align=CENTER)
            c5 = ws.cell(r, 5, credit if credit else None)
            c5.number_format = USD_FMT
            _paint(c5, fill=fill, font=BODY, align=CENTER)
            r += 1
        _extend_filter(ws, led_header, r - 1, 5)
    _extend_filter(ws, 4, 8, 5)


def sheet_sources(wb: Workbook) -> None:
    ws = wb.create_sheet("מקורות")
    _rtl(ws)
    cols = ["סוג", "נתיב / חשבון", "מה נסרק"]
    _widths(ws, [14, 80, 70])
    _banner(ws, 3, "מקורות שנסרקו לבניית התיעוד", "מייל arye.baby@gmail.com, קבצים מקומיים, וואטסאפ, JSON של התוכנה")
    _headers(ws, 4, cols)
    r = 5
    for s in SOURCES:
        fill = ALT if r % 2 == 0 else PAPER
        _paint(ws.cell(r, 1, s["type"]), fill=fill, font=BOLD, align=CENTER)
        _paint(ws.cell(r, 2, s["path"]), fill=fill, font=BODY, align=LEFT)
        _paint(ws.cell(r, 3, s["note"]), fill=fill, font=BODY, align=LEFT)
        ws.row_dimensions[r].height = 22
        r += 1
    _extend_filter(ws, 4, r - 1, 3)

    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    _paint(ws.cell(r, 1, "רשומות הובלה מהתוכנה (shipping_costs.json)"), fill=NAVY, font=WHITE_BOLD, align=LEFT, border=False)
    for col in range(2, 4):
        _paint(ws.cell(r, col), fill=NAVY, border=False)
    r += 1
    cost_headers = ["תאריך", "משלח", "פירוט"]
    _headers(ws, r, cost_headers)
    cost_header = r
    r += 1
    if SHIPPING_COSTS.exists():
        costs = json.loads(SHIPPING_COSTS.read_text(encoding="utf-8"))
        for c in costs:
            fill = ALT if r % 2 == 0 else PAPER
            detail = (
                f"{c.get('rolls_quantity')} גלילים · {c.get('total_weight')} ק״ג · {c.get('cub')} קוב · "
                f"בד {c.get('product_price_usd')}$ · הובלה {c.get('final_cost_incl_domestic')} ₪ · "
                f"PL: {c.get('packing_list') or '—'} · דרישה: {c.get('payment_request') or '—'}"
            )
            _paint(ws.cell(r, 1, c.get("date")), fill=fill, font=BODY, align=CENTER)
            _paint(ws.cell(r, 2, c.get("name")), fill=fill, font=BODY, align=CENTER)
            _paint(ws.cell(r, 3, detail), fill=fill, font=BODY, align=LEFT)
            ws.row_dimensions[r].height = 28
            r += 1
        _extend_filter(ws, cost_header, r - 1, 3)


def sheet_gaps(wb: Workbook) -> None:
    ws = wb.create_sheet("פערים לבדיקה")
    _rtl(ws)
    cols = ["נושא", "פירוט", "סטטוס"]
    _widths(ws, [36, 90, 28])
    _banner(ws, 3, "פערים, כפילויות לא סגורות ונקודות לברור מול ספק/משלח/רואה חשבון", f"{len(GAPS)} פריטים")
    _headers(ws, 4, cols)
    r = 5
    for g in GAPS:
        fill = GOOD if g["status"].startswith("נסגר") else WARN
        if g["status"] == "פתוח":
            fill = BAD
        _paint(ws.cell(r, 1, g["topic"]), fill=fill, font=BOLD, align=LEFT)
        _paint(ws.cell(r, 2, g["detail"]), fill=fill, font=BODY, align=LEFT)
        _paint(ws.cell(r, 3, g["status"]), fill=fill, font=BOLD, align=CENTER)
        ws.row_dimensions[r].height = 48
        r += 1
    _extend_filter(ws, 4, r - 1, 3)


def build() -> Path:
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    sheet_timeline(wb)
    sheet_usd(wb)
    sheet_ils(wb)
    sheet_shipments(wb)
    sheet_invoices(wb)
    sheet_balances(wb)
    sheet_sources(wb)
    sheet_gaps(wb)
    wb.save(OUT_PATH)
    return OUT_PATH


if __name__ == "__main__":
    path = build()
    print(path)
    print("events", len(EVENTS))
    print("usd rows", len(usd_payments()))
    print("ils rows", len(ils_shipping()))
    print("shipments", len(shipments()))
    print("invoices", len(invoices()))
