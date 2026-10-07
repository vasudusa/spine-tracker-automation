# scripts/update_master.py
import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent.parent
MASTER = BASE / "data" / "master.xlsx"
INCOMING = BASE / "data" / "incoming"

def find_latest(folder: Path):
    files = list(folder.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError("No incoming .xlsx files found in data/incoming/")
    return max(files, key=lambda f: f.stat().st_mtime)

def main():
    new_file = find_latest(INCOMING)
    print(f"• Loading new data from: {new_file.name}")
    df_new = pd.read_excel(new_file, sheet_name=0, header=1)  # first sheet

    print(f"• Reading master workbook: {MASTER.name}")
    df_master = pd.read_excel(MASTER, sheet_name=0, header=1)

    # Concatenate and drop duplicates by all columns (or specify subset)
    df_combined = pd.concat([df_master, df_new], ignore_index=True)
    df_combined.drop_duplicates(inplace=True)

    # Write back into master.xlsx
    with pd.ExcelWriter(MASTER, engine="openpyxl", mode="w") as writer:
        df_combined.to_excel(writer, sheet_name="Sheet1", index=False)

    print("✅ Master workbook updated successfully.")

if __name__ == "__main__":
    main()
