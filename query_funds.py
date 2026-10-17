# query_funds.py - run: python query_funds.py
import difflib
from pathlib import Path
import pandas as pd

BASE = Path(__file__).parent
MASTER_PATH = BASE / "mf_master_dataset.csv"
OUT_PATH    = BASE / "filtered_funds_data.csv"

master = pd.read_csv(MASTER_PATH)
master["_key"] = master["Scheme Name"].str.strip().str.lower()

print("Enter fund names separated by commas (partial names work, e.g. 'HDFC Liquid'):")
requested = [n.strip() for n in input("Fund names: ").split(",") if n.strip()]

results, problems = [], []

for name in requested:
    q = name.lower()

    exact = master[master["_key"] == q]
    if not exact.empty:
        results.append(exact)
        continue

    contains = master[master["_key"].str.contains(q, regex=False, na=False)]
    if len(contains) == 1:
        results.append(contains)
    elif len(contains) > 1:
        problems.append((name, f"{len(contains)} matches - be more specific", contains["Scheme Name"].tolist()[:5]))
    else:
        close = difflib.get_close_matches(q, master["_key"].tolist(), n=3, cutoff=0.5)
        sugg = master[master["_key"].isin(close)]["Scheme Name"].tolist()
        problems.append((name, "not found", sugg))

if results:
    out = pd.concat(results).drop(columns="_key").drop_duplicates("Scheme Name")
    out.to_csv(OUT_PATH, index=False)
    print(out.to_string(index=False))
    print(f"\nSaved to {OUT_PATH}")
else:
    print("No funds matched.")

for name, why, sugg in problems:
    print(f"\n'{name}': {why}")
    if sugg:
        print("   Did you mean:", " | ".join(sugg))
