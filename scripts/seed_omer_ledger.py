# -*- coding: utf-8 -*-
"""Seed the supplier ledger (supplier_ledger.json) with Omer / Nibby Baby 2026 entries.

Run once from the project root:  python scripts/seed_omer_ledger.py
Amounts are USD, excluding Turkish VAT (VAT is not tracked; vat_pct = 0).

Sources:
- 01.09.2026 charge  — supplier_documents/4/omer first shipment payment USD 07.09.26.xlsx
                       (14,750 pcs models 200/201/202 = $26,071.24; Omer's statement
                       shows the same invoice as $28,678.36 incl. 10% VAT, which we ignore).
- 03.09.2026 payment — $19,314 advance wired through the forwarder (Express Line),
                       26.8–3.9.2026 (omer first shipment payment USD 07.09.26.xlsx note;
                       confirmed by Arye). Omer's statement (balance 01.10.26.pdf) shows
                       "Nakit 19.250,00" — $64 less, fees/FX. The bank SWIFT of $7,848 is
                       the Baby Basic share of the same advance.
- 30.09.2026 charge  — supplier_documents/4/025-204 invoice.pdf, invoice SF01078
                       (18,278 pcs models 025/204 = $37,383.44 before VAT).
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from optitex_analyzer.core.data_processor import DataProcessor  # noqa: E402

OMER_ID = 4

ENTRIES = [
    {
        "date": "2026-09-01",
        "type": "charge",
        "amount": 26071.24,
        "vat_pct": 0,
        "note": "חשבונית 01.09.2026 — דגמים 200/201/202, 14,750 יח' (משלוח ראשון)",
    },
    {
        "date": "2026-09-03",
        "type": "payment",
        "amount": 19314.00,
        "vat_pct": 0,
        "note": "מקדמה דרך המשלח (Express Line), הועברה 26.8–3.9.2026. בדוח עומר 30.09.2026 נרשם 19,250 — הפרש 64$ עמלות/שער",
    },
    {
        "date": "2026-09-30",
        "type": "charge",
        "amount": 37383.44,
        "vat_pct": 0,
        "note": "חשבונית SF01078 30.09.2026 — דגמים 025/204, 18,278 יח' (משלוח שני)",
    },
]


def main() -> None:
    import os
    os.chdir(ROOT)
    dp = DataProcessor()
    if not dp.get_supplier(OMER_ID):
        raise SystemExit(f"supplier {OMER_ID} not found in suppliers.json")

    existing = dp.get_supplier_ledger(OMER_ID)
    added = 0
    for e in ENTRIES:
        dup = any(
            r.get("date") == e["date"] and r.get("type") == e["type"]
            and abs(float(r.get("amount", 0)) - e["amount"]) < 0.005
            for r in existing
        )
        if dup:
            continue
        dp.add_supplier_ledger_entry(
            OMER_ID, e["type"], e["amount"], date_str=e["date"], vat_pct=e["vat_pct"], note=e["note"],
        )
        added += 1

    bal = dp.get_supplier_balance(OMER_ID)
    print(f"added={added}")
    print(f"charged={bal['charged_net']:,.2f} paid={bal['paid']:,.2f} balance={bal['balance_net']:,.2f}")
    assert abs(bal["balance_net"] - 44140.68) < 0.01, bal["balance_net"]
    print("OK — balance excl. Turkish VAT. (Omer's statement incl. 10% VAT shows Borç 50,550.14 with the advance as 19,250.)")


if __name__ == "__main__":
    main()
