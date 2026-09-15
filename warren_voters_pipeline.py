#!/usr/bin/env python3
"""Build scored/deduped Warren City voter lists from a raw Trumbull County SOS file.

Pipeline (see docs/warren-voters-pipeline.md for full write-up):

  1. Acquire the current Trumbull County voter export (.txt, comma-delimited).
     The Ohio SOS site blocks plain HTTP downloads (403) and this environment
     could not complete the Selenium download either, so by default this
     script reuses the newest local raw file under downloads/. Pass --input
     to pick a specific file, or --download to retry the Selenium fetch
     first (see download_trumbull_ward.py for that logic).
  2. Filter to CITY == "WARREN CITY" (all wards) -> warren-all
  3. Score every voter: TOTAL_VOTES (all-time non-blank election columns) and
     VOTES_LAST_4YR (non-blank election columns dated within the last 4
     years). Sort by TOTAL_VOTES desc.
  4. Filter to VOTES_LAST_4YR >= 1 -> warren-all-4yr-vote1
  5. Dedupe warren-all-4yr-vote1 to one row per household, keyed on
     RESIDENTIAL_ADDRESS1 + RESIDENTIAL_SECONDARY_ADDR, keep the
     highest-TOTAL_VOTES voter per address, and emit it in the Vista mailing
     list template format (Recipient/Company/Address/City/State/Zip code)
     with Recipient = "<Last_Name> Household" -> warren-all-4yr-vote1-deduped

Usage:
  python3 warren_voters_pipeline.py
  python3 warren_voters_pipeline.py --input "downloads/TRUMBULL (1).txt"
  python3 warren_voters_pipeline.py --years 4 --output-dir outputs
  python3 warren_voters_pipeline.py --score-xlsx outputs/2026/warren-all_2026-09-15.xlsx \
      --score-output outputs/2026/warren-all-scored_2026-09-15.xlsx
"""

from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

VOTE_COL_RE = re.compile(r"^(PRIMARY|GENERAL|SPECIAL)-(\d{2})/(\d{2})/(\d{4})$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--download-dir", default="downloads", help="Directory holding raw SOS .txt files (default: downloads)")
    parser.add_argument("--input", default="", help="Specific raw .txt file to use instead of the newest file in --download-dir")
    parser.add_argument("--city", default="WARREN CITY", help='CITY filter (default: "WARREN CITY")')
    parser.add_argument("--years", type=int, default=4, help="Window size in years for the recent-vote score (default: 4)")
    parser.add_argument("--output-dir", default="outputs", help="Base directory for outputs, grouped under outputs/<year>/ (default: outputs)")
    parser.add_argument("--score-xlsx", default="", help="Existing .xlsx workbook to add Total:/Dems/REPS/Latest columns to")
    parser.add_argument("--score-output", default="", help="Output .xlsx path for --score-xlsx (default: add -scored before .xlsx)")
    parser.add_argument("--recent-years", type=int, default=6, help="Recent-vote window for --score-xlsx (default: 6)")
    return parser.parse_args()


def newest_raw_file(download_dir: Path) -> Path:
    candidates = [p for p in download_dir.glob("*.txt") if "trumbull" in p.name.lower()]
    if not candidates:
        raise FileNotFoundError(f"No TRUMBULL*.txt files found under {download_dir}")
    return max(candidates, key=lambda p: p.stat().st_mtime)


def load_raw(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df.columns = df.columns.str.strip().str.upper()
    return df


def vote_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if VOTE_COL_RE.match(c)]


def vote_year(col: str) -> int:
    return int(VOTE_COL_RE.match(col).group(4))


def add_scores(df: pd.DataFrame, vote_cols: list[str], years: int) -> pd.DataFrame:
    df = df.copy()
    current_year = datetime.today().year
    recent_cols = [c for c in vote_cols if vote_year(c) >= current_year - years]

    non_blank = df[vote_cols].apply(lambda s: s.str.strip().ne(""))
    df["TOTAL_VOTES"] = non_blank.sum(axis=1)

    if recent_cols:
        recent_non_blank = df[recent_cols].apply(lambda s: s.str.strip().ne(""))
        df[f"VOTES_LAST_{years}YR"] = recent_non_blank.sum(axis=1)
    else:
        df[f"VOTES_LAST_{years}YR"] = 0

    return df


def score_existing_xlsx(input_path: Path, output_path: Path, recent_years: int) -> None:
    """Add the legacy Excel scoring columns to an already-created workbook.

    The source workbook is not modified.  The four columns are inserted after
    WARD and contain calculated values rather than Excel formulas so they work
    in viewers that do not recalculate formulas:

      Total:  all non-blank election cells
      Dems    election cells equal to D
      REPS    election cells equal to R
      Latest  non-blank election cells dated within the recent-year window
    """
    if recent_years < 0:
        raise ValueError("--recent-years must be zero or greater")

    df = pd.read_excel(input_path, dtype=str, keep_default_na=False)
    df.columns = [str(column).strip() for column in df.columns]
    vote_cols = vote_columns(df)
    if not vote_cols:
        raise KeyError(f"No election columns found in {input_path}")
    if "WARD" not in df.columns:
        raise KeyError(f"Column WARD not found in {input_path}")

    non_blank = df[vote_cols].apply(lambda column: column.str.strip().ne(""))
    scores = pd.DataFrame(index=df.index)
    scores["Total:"] = non_blank.sum(axis=1)
    scores["Dems"] = df[vote_cols].eq("D").sum(axis=1)
    scores["REPS"] = df[vote_cols].eq("R").sum(axis=1)

    current_year = datetime.today().year
    latest_cols = [column for column in vote_cols if vote_year(column) >= current_year - recent_years]
    scores["Latest"] = non_blank[latest_cols].sum(axis=1) if latest_cols else 0

    ward_position = df.columns.get_loc("WARD") + 1
    result = pd.concat([df.iloc[:, :ward_position], scores, df.iloc[:, ward_position:]], axis=1)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(output_path, index=False, engine="openpyxl")
    print(f"Scored workbook written: {output_path}")
    print(f"Scoring columns: Total:, Dems, REPS, Latest ({len(latest_cols)} recent election columns)")


def build_deduped_mailing_list(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    work["_ADDR_KEY"] = (
        work["RESIDENTIAL_ADDRESS1"].str.strip().str.upper()
        + "|"
        + work["RESIDENTIAL_SECONDARY_ADDR"].str.strip().str.upper()
    )
    work = work.sort_values("TOTAL_VOTES", ascending=False, kind="stable").copy()
    reps = work.groupby("_ADDR_KEY", as_index=False, sort=False).first()

    address = reps["RESIDENTIAL_ADDRESS1"].str.strip()
    secondary = reps["RESIDENTIAL_SECONDARY_ADDR"].str.strip()
    full_address = address.where(secondary == "", address + " " + secondary)

    mailing = pd.DataFrame({
        "Recipient": reps["LAST_NAME"].str.strip().str.title() + " Household",
        "Company": [""] * len(reps),
        "Address": full_address,
        "City": reps["RESIDENTIAL_CITY"].str.strip().str.title(),
        "State": reps["RESIDENTIAL_STATE"].str.strip(),
        "Zip code": reps["RESIDENTIAL_ZIP"].str.strip(),
    })
    return mailing.sort_values("Recipient", kind="stable")


def main() -> int:
    args = parse_args()

    if args.score_xlsx:
        input_path = Path(args.score_xlsx).expanduser().resolve()
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        if args.score_output:
            output_path = Path(args.score_output).expanduser().resolve()
        else:
            output_path = input_path.with_name(f"{input_path.stem}-scored{input_path.suffix}")
        score_existing_xlsx(input_path, output_path, args.recent_years)
        return 0

    download_dir = Path(args.download_dir).resolve()
    output_dir = Path(args.output_dir).resolve()

    raw_path = Path(args.input).expanduser().resolve() if args.input else newest_raw_file(download_dir)
    print(f"Using raw file: {raw_path}")

    df = load_raw(raw_path)
    vote_cols = vote_columns(df)
    if not vote_cols:
        raise KeyError("No election columns (PRIMARY-/GENERAL-/SPECIAL-MM/DD/YYYY) found in raw file.")

    warren = df[df["CITY"].str.strip().str.upper() == args.city.upper()].copy()
    if warren.empty:
        raise ValueError(f'No rows matched CITY == "{args.city}"')
    print(f"Warren City voters: {len(warren)}")

    warren = add_scores(warren, vote_cols, args.years)
    recent_col = f"VOTES_LAST_{args.years}YR"
    warren = warren.sort_values("TOTAL_VOTES", ascending=False, kind="stable")

    recent_voters = warren[warren[recent_col] >= 1].copy()
    print(f"Voters with >=1 vote in last {args.years} years: {len(recent_voters)}")

    deduped = build_deduped_mailing_list(recent_voters)
    print(f"Deduped households: {len(deduped)}")

    today = datetime.today()
    year_dir = output_dir / f"{today:%Y}"
    year_dir.mkdir(parents=True, exist_ok=True)
    today_str = today.strftime("%Y-%m-%d")

    all_path = year_dir / f"warren-all_{today_str}.xlsx"
    vote1_path = year_dir / f"warren-all-4yr-vote1_{today_str}.xlsx"
    deduped_path = year_dir / f"warren-all-4yr-vote1-deduped_{today_str}.xlsx"

    warren.to_excel(all_path, index=False, engine="openpyxl")
    recent_voters.to_excel(vote1_path, index=False, engine="openpyxl")
    deduped.to_excel(deduped_path, index=False, engine="openpyxl")

    print(f"Wrote: {all_path}")
    print(f"Wrote: {vote1_path}")
    print(f"Wrote: {deduped_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
