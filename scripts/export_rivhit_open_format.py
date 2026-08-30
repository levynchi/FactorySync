"""Export Rivhit Open Format (מבנה אחיד) ZIP for a tax year.

Calls OpenFormat.ZIP and downloads data.link. Token from RIVHIT_API_TOKEN
or config.json -> rivhit.api_token.
"""
import argparse
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from optitex_analyzer.core.rivhit_api import (  # noqa: E402
    RivhitApiError,
    RivhitOnlineClient,
    default_open_format_year,
)


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
        "חסר טוקן ריווחית. הגדר RIVHIT_API_TOKEN או rivhit.api_token ב-config.json."
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="הפקת קובץ מבנה אחיד (Open Format) מריווחית אונליין",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=default_open_format_year(),
        help="שנת מס (ברירת מחדל: השנה המלאה האחרונה)",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="",
        help="נתיב ZIP לשמירה (ברירת מחדל: exports/open_format/open_format_<year>.zip)",
    )
    args = parser.parse_args()

    dest = Path(args.output) if args.output else (
        ROOT / "exports" / "open_format" / f"open_format_{args.year}.zip"
    )
    client = RivhitOnlineClient(_load_token())
    path = client.download_open_format_zip(args.year, dest)
    print(f"OK: קובץ מבנה אחיד לשנת {args.year} נשמר ב-{path}")


if __name__ == "__main__":
    try:
        main()
    except RivhitApiError as exc:
        raise SystemExit(f"שגיאת ריווחית: {exc}") from exc
