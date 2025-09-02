import json
import shutil
from pathlib import Path

DATA_PATH = Path(__file__).resolve().parents[1] / "fabrics_inventory.json"
BACKUP_PATH = Path(__file__).resolve().parents[1] / "backups"
BACKUP_PATH.mkdir(exist_ok=True)

BACKUP_FILE = BACKUP_PATH / f"fabrics_inventory_backup.json"

TARGET_STATUS = "נשלח"
NEW_LOCATION = "אריה"


def main():
    if not DATA_PATH.exists():
        raise SystemExit(f"Data file not found: {DATA_PATH}")

    # Create a backup
    shutil.copy2(DATA_PATH, BACKUP_FILE)

    with open(DATA_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    updated = 0
    for row in data:
        if str(row.get("status", "")).strip() == TARGET_STATUS:
            if row.get("location") != NEW_LOCATION:
                row["location"] = NEW_LOCATION
                updated += 1

    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Updated {updated} records to location '{NEW_LOCATION}' where status == '{TARGET_STATUS}'.")
    print(f"Backup saved to: {BACKUP_FILE}")


if __name__ == "__main__":
    main()
