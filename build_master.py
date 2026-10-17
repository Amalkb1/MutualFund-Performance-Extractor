# build_master.py - run: python build_master.py
import os, re, glob
from pathlib import Path
import pandas as pd

BASE = Path(__file__).parent
INPUT_DIR   = BASE / "mf_files"
MASTER_PATH = BASE / "mf_master_dataset.csv"
AUM_STD     = "Previous Month Average AUM (Cr.)"   # the final name used in the master

BASE_COLS = [
    "Scheme Name", "Benchmark",
    "Return 1 Year (%) Regular", "Return 1 Year (%) Benchmark",
    "Return 3 Year (%) Regular", "Return 3 Year (%) Benchmark",
    "Return 5 Year (%) Regular", "Return 5 Year (%) Benchmark",
]


def clean_header(h):
    return re.sub(r"\s+", " ", str(h)).strip()


def find_aum_col(headers):
    """Pick whichever AUM column this file has."""
    for name in ("Previous Month Average AUM (Cr.)", "Daily AUM (Cr.)"):
        if name in headers:
            return name
    fallback = [h for h in headers if "AUM" in h]
    return fallback[0] if fallback else None


def extract_file(path):
    raw = pd.read_excel(path, header=None)

    # locate the real header row
    mask = raw.iloc[:, 0].astype(str).str.strip().eq("Scheme Name")
    if not mask.any():
        print(f"  SKIPPED (no 'Scheme Name' header found): {os.path.basename(path)}")
        return None
    hdr_idx = mask.idxmax()

    headers = [clean_header(h) for h in raw.iloc[hdr_idx]]
    df = raw.iloc[hdr_idx + 1:].copy()
    df.columns = headers
    df = df.loc[:, ~df.columns.duplicated()]          # guard against duplicate headers

    aum_src = find_aum_col(list(df.columns))
    keep = BASE_COLS + ([aum_src] if aum_src else [])

    missing = [c for c in keep if c not in df.columns]
    if missing:
        print(f"  WARNING {os.path.basename(path)} missing columns: {missing}")
    df = df.reindex(columns=keep)

    # standardise AUM column name, remember which type it was
    if aum_src:
        df = df.rename(columns={aum_src: AUM_STD})
        df["AUM Type"] = "Daily" if "Daily" in aum_src else "Previous Month Avg"
    else:
        df[AUM_STD] = pd.NA
        df["AUM Type"] = pd.NA

    # drop footnote / blank rows
    df = df.dropna(subset=["Scheme Name", "Benchmark"])

    # numeric conversion
    num_cols = BASE_COLS[2:] + [AUM_STD]
    df[num_cols] = df[num_cols].replace(r"[\^*†‡]+$", "", regex=True)
    df[num_cols] = df[num_cols].apply(pd.to_numeric, errors="coerce")

    # category from file name, e.g. "Liquid Aug.xlsx" -> "Liquid Aug"
    df["Source File"] = os.path.splitext(os.path.basename(path))[0]
    return df


files = sorted(glob.glob(str(INPUT_DIR / "*.xls*")))
print(f"Found {len(files)} file(s)")

frames = []
for f in files:
    d = extract_file(f)
    if d is not None:
        print(f"  OK: {os.path.basename(f)} -> {len(d)} schemes")
        frames.append(d)

master = pd.concat(frames, ignore_index=True)

# remove duplicate schemes (keep the row with the largest AUM)
master["_n"] = master["Scheme Name"].str.strip().str.lower()
master = (master.sort_values(AUM_STD, ascending=False)
                .drop_duplicates("_n")
                .drop(columns="_n")
                .sort_values("Scheme Name")
                .reset_index(drop=True))

master.to_csv(MASTER_PATH, index=False)
print(f"\nMaster dataset saved: {MASTER_PATH}  |  shape = {master.shape}")
print(master.head().to_string())


