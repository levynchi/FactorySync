# -*- coding: utf-8 -*-
"""Build baby_basic_emails.json from scripts/baby_basic_email_data.py.

- Copies attachments found on disk into baby_basic_documents/<thread_id>/ (idempotent).
- Attachments not found locally are kept with path="" (open via Gmail link).
- Preserves manually added items from an existing baby_basic_emails.json
  (threads whose id is not in the data file, plus 'notes' / 'extra_attachments' of known threads).
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from baby_basic_email_data import CATEGORIES, CUSTOMER, THREADS  # noqa: E402

DOCS_ROOT = ROOT / "baby_basic_documents"
OUT_JSON = ROOT / "baby_basic_emails.json"
GMAIL_BASE = "https://mail.google.com/mail/u/0/#all/"


def _rel(p: Path) -> str:
    return str(p.relative_to(ROOT)).replace("\\", "/")


def _copy_attachment(thread_id: str, att: dict) -> dict:
    filename = att["filename"]
    safe_name = filename.replace("/", "_").replace("\\", "_").replace(":", "_")
    dest_dir = DOCS_ROOT / thread_id
    dest = dest_dir / safe_name
    rec = {"filename": filename, "mime": att.get("mime") or "", "path": "", "status": "missing"}
    if dest.exists():
        rec["path"] = _rel(dest)
        rec["status"] = "ok"
        return rec
    for cand in att.get("candidates") or []:
        src = Path(cand)
        if src.is_file():
            if src.resolve().parent == dest_dir.resolve():
                rec["path"] = _rel(src)
                rec["status"] = "ok"
                return rec
            dest_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            rec["path"] = _rel(dest)
            rec["status"] = "ok"
            return rec
    return rec


def build() -> dict:
    existing: dict = {}
    if OUT_JSON.exists():
        try:
            existing = json.loads(OUT_JSON.read_text(encoding="utf-8")) or {}
        except Exception:
            existing = {}
    existing_threads = {t.get("id"): t for t in (existing.get("threads") or []) if t.get("id")}

    threads_out = []
    known_ids = set()
    for t in THREADS:
        tid = t["id"]
        known_ids.add(tid)
        prev = existing_threads.get(tid) or {}
        msgs_out = []
        for m in t.get("messages") or []:
            atts = [_copy_attachment(tid, a) for a in (m.get("attachments") or [])]
            msgs_out.append({
                "date": m["date"],
                "from": m["from"],
                "body": m["body"],
                "attachments": atts,
            })
        gmail_ids = t.get("gmail_ids") or [tid]
        rec = {
            "id": tid,
            "date": t["date"],
            "subject": t["subject"],
            "category": t.get("category") or "אחר",
            "summary": t.get("summary") or "",
            "gmail_ids": gmail_ids,
            "gmail_url": GMAIL_BASE + tid,
            "messages": msgs_out,
            "source": "gmail",
            "notes": prev.get("notes") or "",
            "extra_attachments": prev.get("extra_attachments") or [],
        }
        threads_out.append(rec)

    # keep manually-added threads
    for tid, prev in existing_threads.items():
        if tid not in known_ids and prev.get("source") == "manual":
            threads_out.append(prev)

    threads_out.sort(key=lambda r: r.get("date") or "", reverse=True)

    customer = dict(existing.get("customer") or {})
    customer.update(CUSTOMER)

    return {
        "customer": customer,
        "categories": list(CATEGORIES),
        "threads": threads_out,
    }


def main() -> None:
    data = build()
    OUT_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    total_att = sum(len(m["attachments"]) for t in data["threads"] for m in t.get("messages") or [])
    found = sum(1 for t in data["threads"] for m in t.get("messages") or [] for a in m["attachments"] if a["status"] == "ok")
    print(f"threads={len(data['threads'])} attachments={total_att} found_locally={found}")
    print(f"written: {OUT_JSON}")


if __name__ == "__main__":
    main()
