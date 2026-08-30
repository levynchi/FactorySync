"""Push unit cost 6.48 for Arie flannel long bodysuits to Rivhit.

Updates cost_nis only (does not send sale_nis) for barcodes
7297555020987-7297555021038. Token from RIVHIT_API_TOKEN or config.json.
"""
import importlib.util
import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_api_path = ROOT / "optitex_analyzer" / "core" / "rivhit_api.py"
_spec = importlib.util.spec_from_file_location("rivhit_api", _api_path)
_rivhit_api = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_rivhit_api)
RivhitApiError = _rivhit_api.RivhitApiError
RivhitOnlineClient = _rivhit_api.RivhitOnlineClient

BARCODES = [
    "7297555020987",
    "7297555020994",
    "7297555021007",
    "7297555021014",
    "7297555021021",
    "7297555021038",
]
NEW_COST = 6.48


def _load_token() -> str:
    token = (os.environ.get("RIVHIT_API_TOKEN") or "").strip()
    if token:
        return token
    config_path = ROOT / "config.json"
    if config_path.exists():
        data = json.loads(config_path.read_text(encoding="utf-8"))
        token = str((data.get("rivhit") or {}).get("api_token") or "").strip()
        if token:
            return token
    raise SystemExit(
        "Missing Rivhit API token. Set RIVHIT_API_TOKEN or rivhit.api_token in config.json."
    )


def _item_keys(item: dict) -> set:
    return {
        str(item.get("item_part_num") or "").strip(),
        str(item.get("barcode") or "").strip(),
    }


def main() -> None:
    client = RivhitOnlineClient(_load_token())
    items = client.item_list()
    by_barcode = {}
    for item in items:
        for key in _item_keys(item):
            if key:
                by_barcode.setdefault(key, item)

    updated = []
    missing = []
    for barcode in BARCODES:
        existing = by_barcode.get(barcode)
        if existing is None:
            missing.append(barcode)
            continue
        item_id = existing.get("item_id")
        sale_before = existing.get("sale_nis", existing.get("item_sale_nis"))
        client.item_update(item_id, cost_nis=NEW_COST)
        updated.append(
            {
                "barcode": barcode,
                "item_id": item_id,
                "item_name": existing.get("item_name"),
                "sale_nis": sale_before,
                "cost_nis": NEW_COST,
            }
        )
        print(f"updated {barcode} item_id={item_id} cost={NEW_COST} sale_unchanged={sale_before}")

    if missing:
        raise SystemExit(f"Barcodes not found in Rivhit: {', '.join(missing)}")
    if len(updated) != 6:
        raise SystemExit(f"Expected 6 updates, got {len(updated)}")
    print(f"OK: {len(updated)} items updated to cost {NEW_COST}; sale prices were not sent.")


if __name__ == "__main__":
    try:
        main()
    except RivhitApiError as exc:
        raise SystemExit(f"Rivhit API error: {exc}") from exc
