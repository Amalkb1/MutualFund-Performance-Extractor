"""Core logic: extract fund performance from Mutual Sahi Hai Excel files and manage the master table."""
import re
from pathlib import Path

import pandas as pd

AUM_STD = "Previous Month Average AUM (Cr.)"
BASE_COLS = [
    "Scheme Name", "Benchmark",
    "Return 1 Year (%) Regular", "Return 1 Year (%) Benchmark",
    "Return 3 Year (%) Regular", "Return 3 Year (%) Benchmark",
    "Return 5 Year (%) Regular", "Return 5 Year (%) Benchmark",
]
NUM_COLS = BASE_COLS[2:] + [AUM_STD]
MASTER_COLS = BASE_COLS + [AUM_STD, "AUM Type", "Source File"]

DATA_DIR = Path(__file__).parent / "data"
MASTER_PATH = DATA_DIR / "mf_master_dataset.csv"


def _clean_header(h):
    return re.sub(r"\s+", " ", str(h)).strip()


def _find_aum_col(headers):
    for name in ("Previous Month Average AUM (Cr.)", "Daily AUM (Cr.)"):
        if name in headers:
            return name
    fallback = [h for h in headers if "AUM" in h]
    return fallback[0] if fallback else None


def extract_excel(file, source_name):
    """Return (dataframe, warnings) for one uploaded Excel file. dataframe is None if unusable."""
    warnings = []
    raw = pd.read_excel(file, header=None)

    mask = raw.iloc[:, 0].astype(str).str.strip().eq("Scheme Name")
    if not mask.any():
        return None, ["No 'Scheme Name' header row found - is this the right file?"]
    hdr_idx = mask.idxmax()

    headers = [_clean_header(h) for h in raw.iloc[hdr_idx]]
    df = raw.iloc[hdr_idx + 1:].copy()
    df.columns = headers
    df = df.loc[:, ~df.columns.duplicated()]

    aum_src = _find_aum_col(list(df.columns))
    keep = BASE_COLS + ([aum_src] if aum_src else [])
    missing = [c for c in keep if c not in df.columns]
    if missing:
        warnings.append("Missing columns (left blank): " + ", ".join(missing))
    df = df.reindex(columns=keep)

    if aum_src:
        df = df.rename(columns={aum_src: AUM_STD})
        df["AUM Type"] = "Daily" if "Daily" in aum_src else "Previous Month Avg"
    else:
        df[AUM_STD] = pd.NA
        df["AUM Type"] = pd.NA
        warnings.append("No AUM column found.")

    df = df.dropna(subset=["Scheme Name", "Benchmark"])
    df["Scheme Name"] = df["Scheme Name"].astype(str).str.strip()
    df[NUM_COLS] = df[NUM_COLS].replace(r"[\^*†‡]+$", "", regex=True)
    df[NUM_COLS] = df[NUM_COLS].apply(pd.to_numeric, errors="coerce")
    df["Source File"] = Path(source_name).stem
    return df[MASTER_COLS].reset_index(drop=True), warnings


def load_master():
    if MASTER_PATH.exists():
        return pd.read_csv(MASTER_PATH)
    return pd.DataFrame(columns=MASTER_COLS)


def save_master(df):
    DATA_DIR.mkdir(exist_ok=True)
    df.to_csv(MASTER_PATH, index=False)


def merge_into_master(master, new_frames):
    """Add new data; a scheme that already exists is updated with the newer row."""
    combined = pd.concat([master] + new_frames, ignore_index=True)
    combined["_k"] = combined["Scheme Name"].str.strip().str.lower()
    combined = (combined.drop_duplicates("_k", keep="last")
                        .drop(columns="_k")
                        .sort_values("Scheme Name")
                        .reset_index(drop=True))
    return combined[MASTER_COLS]


def search(master, text="", categories=None):
    """Case-insensitive search; exact scheme names take precedence over partial matches."""
    out = master
    if categories:
        out = out[out["Source File"].isin(categories)]
    terms = [t.strip().lower() for t in text.split(",") if t.strip()]
    if terms:
        names = out["Scheme Name"].str.strip().str.lower()
        mask = pd.Series(False, index=out.index)
        for t in terms:
            exact = names.eq(t)
            if exact.any():
                mask |= exact
                continue

            m = pd.Series(True, index=out.index)
            for word in t.split():
                m &= names.str.contains(word, regex=False, na=False)
            mask |= m
        out = out[mask]
    return out
