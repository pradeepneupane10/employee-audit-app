import subprocess
import sys
import os
import pandas as pd

CHUNKS = [
    ("17 Aug 2026", "23 Aug 2026"),
    ("24 Aug 2026", "31 Aug 2026"),
    ("01 Sep 2026", "07 Sep 2026"),
    ("08 Sep 2026", "10 Sep 2026"),
]

def main():
    print("=" * 80)
    print("      BHADRA MONTH COMPLETE HISTORICAL BACKFILL (CHUNKS 1-4)")
    print("=" * 80)

    for i, (from_dt, to_dt) in enumerate(CHUNKS, 1):
        print(f"\n[CHUNK {i}/{len(CHUNKS)}] Starting scrape for ALL TEAM from {from_dt} to {to_dt}...")
        cmd = [
            sys.executable,
            "audit_automation.py",
            "--employee", "ALL TEAM",
            "--from-date", from_dt,
            "--to-date", to_dt,
            "--non-interactive",
            "--force-full-scrape"
        ]
        res = subprocess.run(cmd)
        if res.returncode != 0:
            print(f"Warning: Chunk {i} exited with code {res.returncode}")
        else:
            print(f"Chunk {i} ({from_dt} - {to_dt}) completed successfully!")

    if os.path.exists("master_audit_history.csv"):
        df = pd.read_csv("master_audit_history.csv")
        print("\n" + "=" * 80)
        print(f"ALL CHUNKS COMPLETE! Total records in master_audit_history.csv: {len(df)}")
        print("=" * 80)

if __name__ == "__main__":
    main()
