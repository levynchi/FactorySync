# -*- coding: utf-8 -*-
"""Build import_history.json and copy external Turkey-import files into the project.

Idempotent: existing copies under import_documents/<id>/ are skipped.
In-project files (supplier_documents, packing_lists, payment_requests) are
linked in place and not copied.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from turkey_timeline_data import AKDEM_LEDGER_SUMMARY, EVENTS  # noqa: E402

HOME = Path.home()
BILAL = HOME / "OneDrive" / "שולחן העבודה" / "ביללאל"
HAMOUDI = HOME / "OneDrive" / "שולחן העבודה" / "חמודי"
DOWNLOADS = HOME / "Downloads"

DOCS_ROOT = ROOT / "import_documents"
OUT_JSON = ROOT / "import_history.json"

IN_PROJECT_PREFIXES = (
    "supplier_documents/",
    "supplier_documents\\",
    "packing_lists/",
    "packing_lists\\",
    "payment_requests/",
    "payment_requests\\",
    "exports/",
    "exports\\",
)

GMAIL_HEX = re.compile(r"\b([0-9a-f]{16})\b", re.I)
GMAIL_SHORT = re.compile(r"Gmail\s+([0-9a-f]{8,15})\b", re.I)
SUPPLIER_DOC = re.compile(r"supplier_documents[/\\][^\s,+]+")


def P(*parts: str) -> Path:
    return Path(*parts)


# (date, kind, party_or_None) -> [(label, path)]
# party is used only when the same date+kind appears twice.
ATTACHMENTS: dict[tuple, list[tuple[str, Path]]] = {
    ("2022-08-04", "משלוח יצא", None): [
        ("Packing list 04.08.2022", BILAL / "03_רשימות_משקל_Packing" / "04.08.2022 - Arye Moshe Levy Çeki Listesi.xlsx"),
    ],
    ("2022-08-16", "פרופורמה", None): [
        ("פרופורמת Stilteks 16.08.2022", ROOT / "supplier_documents" / "5" / "Stilteks_proforma_16.08.2022.pdf"),
    ],
    ("2022-08-23", "חשבונית", None): [
        ("חשבונית 23.08.2022", BILAL / "02_חשבוניות_ופרופורמות" / "23.08.2022 - Arye Moshe Invoice.pdf"),
        ("Packing list 23.08.2022", BILAL / "03_רשימות_משקל_Packing" / "23.08.2022 - Arye Moshe Levy Çeki Listesi.xlsx"),
    ],
    ("2022-08-31", "תשלום ספק", None): [
        ("SWIFT Stilteks 31.08.2022", ROOT / "supplier_documents" / "5" / "Stilteks_SWIFT_31.08.2022.pdf"),
        ("CCF 31.08.2022", BILAL / "04_SWIFT_ובנק" / "CCF31082022_0001.pdf"),
    ],
    ("2022-09-01", "משלוח יצא", None): [
        ("YK105925 TELEX", BILAL / "05_שילוח_Telex_HBL" / "YK105925 TELEX_000638.pdf"),
    ],
    ("2022-10-12", "משלוח יצא", None): [
        ("Packing list 12.10.2022 KALANLAR", BILAL / "03_רשימות_משקל_Packing" / "12.10.2022 - Arye Moshe Levy Çeki Listesi KALANLAR.xlsx"),
    ],
    ("2022-12-05", "תשלום ספק", None): [
        ("SWIFT 5,000$ 5.12.2022", BILAL / "04_SWIFT_ובנק" / "swift for 5000 dollar 5.12.2022.pdf"),
    ],
    ("2023-02-22", "תשלום ספק", None): [
        ("CCF 23.02.2023", BILAL / "01_שנת_2023" / "CCF23022023.pdf"),
    ],
    ("2023-02-22", "חשבונית", None): [
        ("חשבונית מסחרית Stilteks 22.02.2023", ROOT / "supplier_documents" / "5" / "Stilteks_commercial_invoice_22.02.2023.pdf"),
        ("Packing list Stilteks 22.02.2023", ROOT / "supplier_documents" / "5" / "Stilteks_packing_list_22.02.2023.pdf"),
        ("פרופורמת Stilteks 22.02.2023", ROOT / "supplier_documents" / "5" / "Stilteks_proforma_22.02.2023.pdf"),
        ("חשבונית Arye 22.02.2023", BILAL / "01_שנת_2023" / "22.02.2023 - Arye Moshe Invoice.pdf"),
    ],
    ("2023-02-22", "משלוח יצא", None): [
        ("Packing list 22.02.2023 Arye Moshe Levy", BILAL / "01_שנת_2023" / "22.02.2023 Arye Moshe Levy Çeki Listesi.xlsx"),
        ("Packing list Stilteks 22.02.2023", ROOT / "supplier_documents" / "5" / "Stilteks_packing_list_22.02.2023.pdf"),
    ],
    ("2023-08-07", "חשבונית", None): [
        ("חשבונית 07.08.2023", BILAL / "01_שנת_2023" / "07.08.2023 - Arye Moshe Invoice.pdf"),
        ("Çeki Listesi 07.08.2023", BILAL / "01_שנת_2023" / "07.08.2023 Arye Moshe Levy Çeki Listesi.xlsx"),
    ],
    ("2023-10-19", "חשבונית", None): [
        ("חשבונית 19.10.2023", BILAL / "01_שנת_2023" / "19.10.2023 - Arye Moshe Invoice.pdf"),
    ],
    ("2024-01-30", "תשלום ספק", None): [
        ("SWIFT 5,000$ 30.01.2024", DOWNLOADS / "swift 30.01.24 akdem.pdf"),
    ],
    ("2024-02-05", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2024-05-28", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
        ("כרטסת Adel", BILAL / "00_כרטסות_ובאלנסים" / "ADEL.xlsx"),
    ],
    ("2024-05-28", "משלוח יצא", None): [
        ("Check List 28.05.2024", DOWNLOADS / "Arye Moshe Check List 28.05.2024.xlsx"),
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-06-20", "תשלום שילוח/עמילות", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-07-08", "תשלום סחורה דרך משלח", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2024-07-25", "משלוח יצא", None): [
        ("Packing list 25.07.2024", DOWNLOADS / "Arye Moshe Packing List 25.07.2024.xlsx"),
        ("CONT 3091 01.08.2024", HAMOUDI / "3091  CONT 01.08.2024.pdf"),
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-07-30", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2024-08-14", "תשלום שילוח/עמילות", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-08-15", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2024-09-04", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-09-04", "משלוח הגיע", None): [
        ("Check List 04.09.2024", DOWNLOADS / "Arye Moshe Check List 04.09.2024.xlsx"),
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-10-28", "תשלום סחורה דרך משלח", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-11-09", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2024-11-26", "משלוח הגיע", None): [
        ("AKDEM 26.11.24", DOWNLOADS / "AKDEM 26.11.24 EDIT FOR AIR TABLE.xlsx"),
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2024-12-02", "תשלום סחורה דרך משלח", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2025-02-06", "תשלום סחורה דרך משלח", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
        ("CamScanner 22.02.2025", HAMOUDI / "CamScanner 22-02-2025 12.53.pdf"),
    ],
    ("2025-02-20", "תשלום סחורה דרך משלח", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
        ("CamScanner 27.02.2025", HAMOUDI / "CamScanner 27-02-2025 13.29.pdf"),
    ],
    ("2025-03-03", "מסמך", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2025-03-11", "פרופורמה", None): [
        ("ADEL GROUP פרופורמה", HAMOUDI / "ADEL GROUP İSRAİL MOSHE PROFORMA.pdf"),
    ],
    ("2025-03-11", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2025-03-12", "תשלום סחורה דרך משלח", None): [
        ("אדל גרופ חשבוניות מול תשלומים", HAMOUDI / "אדל_גרופ_חשבוניות_מול_תשלומים.xlsx"),
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2025-03-25", "משלוח הגיע", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2025-04-03", "תשלום סחורה דרך משלח", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2025-04-07", "פרופורמה", None): [
        ("3091.pdf", HAMOUDI / "3091.pdf"),
    ],
    ("2025-04-15", "תשלום סחורה דרך משלח", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2025-04-22", "משלוח הגיע", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
        ("CamScanner 30.04.2025", HAMOUDI / "CamScanner 30-04-2025 11.18.pdf"),
    ],
    ("2025-04-30", "תשלום שילוח/עמילות", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2025-05-12", "תשלום שילוח/עמילות", None): [
        ("סיכום משלוחים חמודי", HAMOUDI / "סיכום_משלוחים_והעברות_סחורה.xlsx"),
    ],
    ("2025-07-07", "תשלום ספק", None): [
        ("SWIFT 7,582$ 07.2025", DOWNLOADS / "swift.pdf"),
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2025-08-20", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2025-08-27", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
        ("akdem 07.09.25.xls", DOWNLOADS / "akdem 07.09.25.xls"),
    ],
    ("2025-10-12", "תשלום ספק", None): [
        ("CCF 15.10.2025", BILAL / "04_SWIFT_ובנק" / "CCF15102025.pdf"),
    ],
    ("2025-10-22", "תשלום סחורה דרך משלח", None): [
        ("BANK DETAILS VAKIFBANK", BILAL / "04_SWIFT_ובנק" / "BANK DETAILS VAKIFBANK.pdf"),
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2025-11-04", "מסמך", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2026-02-24", "משלוח יצא", None): [
        ("Arye Packing List 24.02.2026", DOWNLOADS / "Arye Packing List.xlsx"),
    ],
    ("2026-03-30", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2026-04-09", "תשלום שילוח/עמילות", None): [
        ("חישוב קוב אקספרסליין 2026", ROOT / "import_documents" / "אקספרסליין_חישוב_קוב_2026.xlsx"),
    ],
    ("2026-04-21", "מסמך", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2026-04-27", "חשבונית", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2026-05-10", "תשלום שילוח/עמילות", None): [
        ("דרישת תשלום #1581 (10.05.2026) — 40$ לקרטון (מסמך מקורי)", ROOT / "payment_requests" / "דרישת תשלום #1581 מאת אקספרסליין קרגו בעיימ 40 דולר לקרטון_20260510.pdf"),
        ("Packing list Bilal — ARYE ÇEKİ LİSTESİ (14.5.2026)", ROOT / "import_documents" / "125" / "ARYE ÇEKİ LİSTESİ.xls"),
        ("חישוב קוב אקספרסליין 2026", ROOT / "import_documents" / "אקספרסליין_חישוב_קוב_2026.xlsx"),
    ],
    ("2026-06-10", "משלוח יצא", None): [
        ("ARYE ÇEKİ LİSTESİ 10.6.2026", DOWNLOADS / "ARYE ÇEKİ LİSTESİ 10.6.2026.xls"),
        ("ARYE ÇEKİ LİSTESİ.xls", DOWNLOADS / "ARYE ÇEKİ LİSTESİ.xls"),
    ],
    ("2026-06-16", "מסמך", None): [
        ("כרטסת Akdem", BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"),
    ],
    ("2026-08-26", "תשלום סחורה דרך משלח", None): [
        ("SWIFT 7,848$ Express Line", DOWNLOADS / "LEVI MOSHE ARIE - SWIFT WIRE TRANSFER.pdf"),
    ],
    ("2026-09-01", "משלוח יצא", None): [
        ("תשלום משלוח ראשון עומר", ROOT / "supplier_documents" / "4" / "omer first shipment payment USD 07.09.26.xlsx"),
        ("first shipment omer 07.09.26", ROOT / "supplier_documents" / "4" / "first shipment omer 07.09.26 .xlsb.xlsx"),
    ],
    ("2026-09-07", "חשבונית", None): [
        ("תשלום משלוח ראשון עומר", ROOT / "supplier_documents" / "4" / "omer first shipment payment USD 07.09.26.xlsx"),
        ("first shipment omer 07.09.26", ROOT / "supplier_documents" / "4" / "first shipment omer 07.09.26 .xlsb.xlsx"),
        ("הזמנת בגדים 200-201-202", ROOT / "supplier_documents" / "4" / "200-201-202.xlsx"),
    ],
    ("2026-09-09", "תשלום שילוח/עמילות", None): [
        ("חישוב קוב אקספרסליין 2026", ROOT / "import_documents" / "אקספרסליין_חישוב_קוב_2026.xlsx"),
    ],
    ("2026-09-14", "מסמך", None): [
        ("פתק חישוב ליניב 14.09.2026 — הפרש 2,409 ₪", ROOT / "import_documents" / "חישוב_הפרש_יניב_14.09.2026.jpg"),
        ("חישוב קוב אקספרסליין 2026", ROOT / "import_documents" / "אקספרסליין_חישוב_קוב_2026.xlsx"),
    ],
}


def _is_in_project(path: Path) -> bool:
    try:
        path.resolve().relative_to(ROOT.resolve())
        return True
    except ValueError:
        return False


def _rel(path: Path) -> str:
    try:
        return path.resolve().relative_to(ROOT.resolve()).as_posix()
    except ValueError:
        return path.as_posix()


def _lookup_manual(event: dict) -> list[tuple[str, Path]]:
    date, kind, party = event["date"], event["kind"], event["party"]
    for key in ((date, kind, party), (date, kind, None)):
        if key in ATTACHMENTS:
            return list(ATTACHMENTS[key])
    return []


def _glob_dir(folder: Path, *needles: str) -> list[Path]:
    if not folder.is_dir():
        return []
    found: list[Path] = []
    for needle in needles:
        if not needle:
            continue
        needle_l = needle.lower()
        for p in folder.iterdir():
            if p.is_file() and needle_l in p.name.lower():
                found.append(p)
    # Prefer one canonical file per stem (latest mtime among near-duplicates)
    by_stem: dict[str, Path] = {}
    for p in found:
        stem = re.sub(r"_\d{8}_\d{6}$", "", p.stem)
        prev = by_stem.get(stem)
        if prev is None or p.stat().st_mtime > prev.stat().st_mtime:
            by_stem[stem] = p
    return list(by_stem.values())


def _from_source_string(event: dict) -> list[tuple[str, Path]]:
    src = event.get("source") or ""
    out: list[tuple[str, Path]] = []
    for m in SUPPLIER_DOC.finditer(src):
        rel = m.group(0).replace("\\", "/")
        # trim trailing punctuation
        rel = rel.rstrip(".,;:+")
        p = ROOT / rel
        if p.is_file():
            out.append((p.name, p))
        else:
            # "supplier_documents/5 packing list" is not a real path
            parent = ROOT / Path(rel).parent
            if parent.is_dir() and "packing" in src.lower():
                for extra in parent.glob("*packing*"):
                    if extra.is_file():
                        out.append((extra.name, extra))
    if "packing_lists" in src or "packing list" in src.lower() or "Çeki" in src or "ceki" in src.lower():
        needles = []
        for token in ("20.08.2025", "30.03.2026", "arye new", "Arye"):
            if token.lower() in src.lower() or token in src:
                needles.append(token)
        if "20.08.2025" in src:
            needles = ["20.08.2025"]
        if "30.03.2026" in src:
            needles = ["30.03.2026"]
        if "arye new" in src.lower():
            needles = ["arye new"]
        for p in _glob_dir(ROOT / "packing_lists", *needles):
            out.append((p.name, p))
    if "payment_requests" in src:
        needles = []
        for token in ("אריה 783", "אריה 816", "דרישת תשלום_20260409", "#2432", "#1581"):
            if token in src:
                needles.append(token)
        if not needles:
            needles = ["דרישת"]
        for p in _glob_dir(ROOT / "payment_requests", *needles):
            out.append((p.name, p))
    if "ARYE MOSHE balance.xlsx" in src or "כרטסת" in src:
        p = BILAL / "00_כרטסות_ובאלנסים" / "ARYE MOSHE balance.xlsx"
        if p.is_file():
            out.append(("כרטסת Akdem", p))
    return out


def _from_shipping_costs(event: dict) -> list[tuple[str, Path]]:
    """Attach packing_list / payment_request from shipping_costs.json when dates match."""
    costs_path = ROOT / "shipping_costs.json"
    if not costs_path.is_file():
        return []
    try:
        records = json.loads(costs_path.read_text(encoding="utf-8"))
    except Exception:
        return []
    y, m, d = event["date"].split("-")
    candidates = {
        f"{int(d)}/{int(m)}/{y}",
        f"{d}/{m}/{y}",
        f"{int(d):02d}/{int(m):02d}/{y}",
    }
    out: list[tuple[str, Path]] = []
    if event["kind"] not in ("תשלום שילוח/עמילות", "משלוח יצא", "משלוח הגיע", "חשבונית"):
        return out
    for rec in records:
        if rec.get("date") not in candidates:
            continue
        pl = rec.get("packing_list") or ""
        pr = rec.get("payment_request") or ""
        if pl:
            p = ROOT / "packing_lists" / pl
            if p.is_file():
                out.append((f"Packing list — {pl}", p))
        if pr:
            p = ROOT / "payment_requests" / pr
            if p.is_file():
                out.append((f"דרישת תשלום — {pr}", p))
    return out


def _gmail_threads(source: str) -> list[dict]:
    seen: set[str] = set()
    threads: list[dict] = []
    if "..." in source:
        source = re.sub(r"[0-9a-f]{4,}\.\.\.", "", source, flags=re.I)
    for rx in (GMAIL_HEX, GMAIL_SHORT):
        for m in rx.finditer(source):
            gid = m.group(1).lower()
            if gid in seen or gid in ("packing",):
                continue
            seen.add(gid)
            threads.append({"id": gid, "label": f"Gmail {gid}"})
    return threads


def _place_attachment(event_id: int, label: str, src: Path, copied: list, missing: list, linked: list) -> dict | None:
    if not src.exists() or not src.is_file():
        missing.append(str(src))
        return None
    if _is_in_project(src):
        rel = _rel(src)
        # If it already lives under this event's import_documents folder, keep it
        linked.append(rel)
        return {"label": label, "path": rel}
    dest_dir = DOCS_ROOT / str(event_id)
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    if not dest.exists():
        shutil.copy2(src, dest)
        copied.append(str(dest))
    else:
        linked.append(_rel(dest))
    return {"label": label, "path": _rel(dest)}


def _dedupe(items: list[tuple[str, Path]]) -> list[tuple[str, Path]]:
    seen: set[str] = set()
    out: list[tuple[str, Path]] = []
    for label, path in items:
        try:
            key = str(path.resolve()).lower()
        except OSError:
            key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        out.append((label, path))
    return out


def build() -> Path:
    os.chdir(ROOT)
    events_sorted = sorted(
        EVENTS,
        key=lambda e: (e.get("date") or "", e.get("kind") or "", e.get("party") or ""),
    )
    copied: list[str] = []
    missing: list[str] = []
    linked: list[str] = []
    rows: list[dict] = []

    for i, src_event in enumerate(events_sorted, start=1):
        event = dict(src_event)
        event["id"] = i
        candidates = _dedupe(
            _lookup_manual(event) + _from_source_string(event) + _from_shipping_costs(event)
        )
        attachments: list[dict] = []
        for label, path in candidates:
            rec = _place_attachment(i, label, path, copied, missing, linked)
            if rec:
                # de-dupe by stored path
                if any(a["path"] == rec["path"] for a in attachments):
                    continue
                attachments.append(rec)
        event["attachments"] = attachments
        event["gmail_threads"] = _gmail_threads(event.get("source") or "")
        rows.append(event)

    payload = {
        "akdem_claimed_balance_usd": AKDEM_LEDGER_SUMMARY.get("claimed_balance_usd"),
        "akdem_ledger": AKDEM_LEDGER_SUMMARY,
        "events": rows,
    }
    OUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    missing_unique = sorted(set(missing))
    print(f"events: {len(rows)}")
    print(f"copied: {len(copied)}")
    print(f"linked in-project: {len(linked)}")
    print(f"missing sources: {len(missing_unique)}")
    with_files = sum(1 for e in rows if e["attachments"])
    print(f"events with attachments: {with_files}")
    with_mail = sum(1 for e in rows if e["gmail_threads"])
    print(f"events with gmail threads: {with_mail}")
    if missing_unique:
        print("--- missing ---")
        for m in missing_unique:
            print(m)
    print(f"wrote {OUT_JSON}")
    return OUT_JSON


if __name__ == "__main__":
    build()
