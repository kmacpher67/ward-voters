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
  3. Score every voter: Local_Tot (odd-year non-blank election columns) and
     VOTES_LAST_4YR (non-blank election columns dated within the last 4
     years). Sort by Local_Tot desc.
  4. Filter to VOTES_LAST_4YR >= 1 -> warren-all-4yr-vote1
  5. Dedupe warren-all-4yr-vote1 to one row per household, keyed on
     RESIDENTIAL_ADDRESS1 + RESIDENTIAL_SECONDARY_ADDR, rank by recent local
     votes then lifetime votes, and emit a Vista mailing-list CSV with the
     recent-vote score retained as the final column for trimming decisions.

Usage:
  python3 warren_voters_pipeline.py
  python3 warren_voters_pipeline.py --input "downloads/TRUMBULL (1).txt"
  python3 warren_voters_pipeline.py --years 4 --output-dir outputs
  python3 warren_voters_pipeline.py --score-xlsx outputs/2026/warren-all_2026-09-15.xlsx \
      --score-output outputs/2026/warren-all-scored_2026-09-15.xlsx
  python3 warren_voters_pipeline.py --ward-xlsx outputs/2026/warren-all-scored_2026-09-15.xlsx \
      --ward 4 --ward-output outputs/2026/warren-ward4-scored_2026-09-15.xlsx
"""

from __future__ import annotations

import argparse
import csv
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

VOTE_COL_RE = re.compile(r"^(PRIMARY|GENERAL|SPECIAL)-(\d{2})/(\d{2})/(\d{4})$")
LOCAL_TOTAL_COL = "Local_Tot"
LEGACY_TOTAL_COL = "TOTAL_VOTES"
SCORE_VALUE_COLUMNS = ("Total:", "Dems", "REPS", "Latest", LOCAL_TOTAL_COL, LEGACY_TOTAL_COL, "VOTES_LAST_4YR")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--download-dir", default="downloads", help="Directory holding raw SOS .txt files (default: downloads)")
    parser.add_argument("--input", default="", help="Specific raw .txt file to use instead of the newest file in --download-dir")
    parser.add_argument("--city", default="WARREN CITY", help='CITY filter (default: "WARREN CITY")')
    parser.add_argument("--years", type=int, default=4, help="Window size in years for the recent-vote score (default: 4)")
    parser.add_argument("--output-dir", default="outputs", help="Base directory for outputs, grouped under outputs/<year>/ (default: outputs)")
    parser.add_argument(
        "--exceptions-csv",
        default="config/warren_no_delivery_addresses.csv",
        help="CSV of address-level no-delivery exceptions (default: config/warren_no_delivery_addresses.csv)",
    )
    parser.add_argument(
        "--manual-addresses-csv",
        default="config/warren_manual_addresses.csv",
        help="CSV of manually requested mailing addresses (default: config/warren_manual_addresses.csv)",
    )
    parser.add_argument(
        "--max-addresses",
        type=int,
        default=3000,
        help="Maximum number of deduped mailing addresses; 0 means unlimited (default: 3000)",
    )
    parser.add_argument(
        "--min-recent-votes",
        type=int,
        default=1,
        help="Minimum recent-vote score before address deduplication (default: 1)",
    )
    parser.add_argument(
        "--verify-exceptions",
        action="store_true",
        help="Verify no-delivery addresses against the raw voter file and write an audit CSV, then exit",
    )
    parser.add_argument(
        "--clean-exceptions",
        action="store_true",
        help="With --verify-exceptions, deactivate definitive elected-official address mismatches in the exceptions CSV",
    )
    parser.add_argument("--score-xlsx", default="", help="Existing .xlsx workbook to add Total:/Dems/REPS/Latest columns to")
    parser.add_argument("--score-output", default="", help="Output .xlsx path for --score-xlsx (default: add -scored before .xlsx)")
    parser.add_argument(
        "--score-format",
        choices=("values", "formulas"),
        default="formulas",
        help="Write --score-xlsx inserted columns as live Excel formulas or calculated numbers (default: formulas)",
    )
    parser.add_argument("--ward-xlsx", default="", help="Existing .xlsx workbook to filter down to a single WARD (e.g. an already-scored workbook)")
    parser.add_argument("--ward", default="", help='Ward to keep for --ward-xlsx, e.g. "4" or "WARREN-WARD 4"')
    parser.add_argument("--ward-output", default="", help="Output .xlsx path for --ward-xlsx (default: add -ward<N> before .xlsx)")
    parser.add_argument("--recent-years", type=int, default=6, help="Recent-vote window for --score-xlsx (default: 6)")
    parser.add_argument(
        "--include-presidential-general",
        dest="exclude_presidential_general",
        action="store_false",
        default=True,
        help="Count presidential-year (year %% 4 == 0) GENERAL elections in the recent/latest vote score "
        "(default: excluded, since presidential-year turnout is not representative of local-election turnout)",
    )
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


def is_odd_year_vote(col: str) -> bool:
    """True for odd-year elections, the Ohio local/municipal cycle used here."""
    return vote_year(col) % 2 == 1


def is_presidential_general(col: str) -> bool:
    """True for a GENERAL election column in a presidential year (year % 4 == 0)."""
    match = VOTE_COL_RE.match(col)
    return match.group(1) == "GENERAL" and int(match.group(4)) % 4 == 0


def add_scores(df: pd.DataFrame, vote_cols: list[str], years: int, exclude_presidential_general: bool = True) -> pd.DataFrame:
    df = df.copy()
    current_year = datetime.today().year
    recent_cols = [c for c in vote_cols if vote_year(c) >= current_year - years]
    if exclude_presidential_general:
        recent_cols = [c for c in recent_cols if not is_presidential_general(c)]

    non_blank = df[vote_cols].apply(lambda s: s.str.strip().ne(""))
    local_cols = [c for c in vote_cols if is_odd_year_vote(c)]
    df[LOCAL_TOTAL_COL] = non_blank[local_cols].sum(axis=1) if local_cols else 0

    if recent_cols:
        recent_non_blank = df[recent_cols].apply(lambda s: s.str.strip().ne(""))
        df[f"VOTES_LAST_{years}YR"] = recent_non_blank.sum(axis=1)
    else:
        df[f"VOTES_LAST_{years}YR"] = 0

    return df


def excel_column_letter(index: int) -> str:
    """Convert a 1-based column index to an Excel column letter."""
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def nonblank_count_formula(row: int, columns: list[int]) -> str:
    """Build an Excel formula counting nonblank cells in non-contiguous columns."""
    if not columns:
        return "=0"
    parts = [f'IF({excel_column_letter(column)}{row}<>"",1,0)' for column in columns]
    return f"={'+'.join(parts)}"


def build_summary_sheet(wb, ws, header_to_col: dict[str, int], max_row: int, column_maxes: dict[str, int]) -> None:
    """Add/replace a 'Summary' sheet with a COUNTIF super-voter ladder.

    For each score column present, one row per threshold from >=1 up to that
    column's own max value in the data, e.g. row ">=3" for Dems shows
    =COUNTIF(Dems range, ">="&3). Columns with a lower max simply show 0 on
    the rows above their own max, so the table stays rectangular.
    """
    overall_max = max(column_maxes.values(), default=0)
    if overall_max <= 0:
        return
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    summary = wb.create_sheet("Summary")

    sheet_ref = f"'{ws.title}'" if any(ch in ws.title for ch in " -") else ws.title
    summary.cell(row=1, column=1, value="Threshold (>=)")
    ordered_headers = [header for header in header_to_col if header in column_maxes]
    for col_offset, header in enumerate(ordered_headers, start=2):
        summary.cell(row=1, column=col_offset, value=header)
        col_letter = excel_column_letter(header_to_col[header])
        data_range = f"{sheet_ref}!${col_letter}$2:${col_letter}${max_row}"
        for threshold in range(1, overall_max + 1):
            row = threshold + 1
            summary.cell(row=row, column=1, value=threshold)
            summary.cell(row=row, column=col_offset, value=f'=COUNTIF({data_range},">="&$A{row})')


def score_existing_xlsx_values(input_path: Path, output_path: Path, recent_years: int, exclude_presidential_general: bool = True) -> int:
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
    if exclude_presidential_general:
        latest_cols = [column for column in latest_cols if not is_presidential_general(column)]
    scores["Latest"] = non_blank[latest_cols].sum(axis=1) if latest_cols else 0

    local_cols = [column for column in vote_cols if is_odd_year_vote(column)]
    df[LOCAL_TOTAL_COL] = non_blank[local_cols].sum(axis=1) if local_cols else 0
    if LEGACY_TOTAL_COL in df.columns:
        df = df.drop(columns=[LEGACY_TOTAL_COL])
    existing_score_cols = [LOCAL_TOTAL_COL]
    existing_score_cols.extend(column for column in df.columns if re.match(r"^VOTES_LAST_\d+YR$", column))
    for column in existing_score_cols:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="raise")

    ward_position = df.columns.get_loc("WARD") + 1
    after_ward = df.iloc[:, ward_position:]
    moved_score_cols = [column for column in existing_score_cols if column in after_ward.columns]
    remaining_after_ward = after_ward.drop(columns=moved_score_cols)
    result = pd.concat(
        [df.iloc[:, :ward_position], scores, df[moved_score_cols], remaining_after_ward],
        axis=1,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(output_path, index=False, engine="openpyxl")

    from openpyxl import load_workbook as _load_workbook

    summary_headers = ["Total:", "Dems", "REPS", "Latest"] + moved_score_cols
    column_maxes = {
        header: int(result[header].max())
        for header in summary_headers
        if header in result.columns and len(result) and result[header].notna().any()
    }
    header_to_col = {header: result.columns.get_loc(header) + 1 for header in column_maxes}
    wb = _load_workbook(output_path)
    build_summary_sheet(wb, wb.active, header_to_col, len(result) + 1, column_maxes)
    wb.save(output_path)

    return len(latest_cols)


def score_existing_xlsx_formulas(input_path: Path, output_path: Path, recent_years: int, exclude_presidential_general: bool = True) -> int:
    from openpyxl import load_workbook

    wb = load_workbook(input_path)
    ws = wb.active
    headers = {str(ws.cell(1, column).value or "").strip(): column for column in range(1, ws.max_column + 1)}
    carried_score_headers = [
        header for header in headers
        if header == LOCAL_TOTAL_COL or header == LEGACY_TOTAL_COL or re.match(r"^VOTES_LAST_\d+YR$", header)
    ]
    for column in sorted((headers[header] for header in carried_score_headers), reverse=True):
        ws.delete_cols(column)

    headers = {str(ws.cell(1, column).value or "").strip(): column for column in range(1, ws.max_column + 1)}
    if "WARD" not in headers:
        raise KeyError(f"Column WARD not found in {input_path}")
    original_vote_columns = {
        header: column for header, column in headers.items()
        if VOTE_COL_RE.match(header)
    }
    if not original_vote_columns:
        raise KeyError(f"No election columns found in {input_path}")

    extra_headers = [LOCAL_TOTAL_COL]
    extra_headers.extend(header for header in carried_score_headers if re.match(r"^VOTES_LAST_\d+YR$", header))

    insert_position = headers["WARD"] + 1
    insert_amount = 4 + len(extra_headers)
    ws.insert_cols(insert_position, amount=insert_amount)
    for offset, header in enumerate(("Total:", "Dems", "REPS", "Latest")):
        ws.cell(row=1, column=insert_position + offset, value=header)
    for offset, header in enumerate(extra_headers, start=4):
        ws.cell(row=1, column=insert_position + offset, value=header)

    inserted_before_votes = {
        header: column + insert_amount if column >= insert_position else column
        for header, column in original_vote_columns.items()
    }
    vote_positions = sorted(inserted_before_votes.values())
    first_vote_letter = excel_column_letter(vote_positions[0])
    last_vote_letter = excel_column_letter(vote_positions[-1])

    current_year = datetime.today().year
    latest_cols = [
        column for column in original_vote_columns
        if vote_year(column) >= current_year - recent_years
    ]
    if exclude_presidential_general:
        latest_cols = [column for column in latest_cols if not is_presidential_general(column)]
    latest_positions = [inserted_before_votes[column] for column in latest_cols]
    local_positions = [
        inserted_before_votes[column] for column in original_vote_columns
        if is_odd_year_vote(column)
    ]
    recent_score_positions = {}
    for header in extra_headers:
        match = re.match(r"^VOTES_LAST_(\d+)YR$", header)
        if not match:
            continue
        years = int(match.group(1))
        columns = [column for column in original_vote_columns if vote_year(column) >= current_year - years]
        if exclude_presidential_general:
            columns = [column for column in columns if not is_presidential_general(column)]
        recent_score_positions[header] = [inserted_before_votes[column] for column in columns]

    column_maxes = {header: 0 for header in ("Total:", "Dems", "REPS", "Latest", *extra_headers)}
    for row in range(2, ws.max_row + 1):
        vote_range = f"${first_vote_letter}${row}:${last_vote_letter}${row}"
        ws.cell(row=row, column=insert_position, value=f"=COUNTA({vote_range})")
        ws.cell(row=row, column=insert_position + 1, value=f'=COUNTIF({vote_range},"D")')
        ws.cell(row=row, column=insert_position + 2, value=f'=COUNTIF({vote_range},"R")')
        ws.cell(row=row, column=insert_position + 3, value=nonblank_count_formula(row, latest_positions))
        for offset, header in enumerate(extra_headers, start=4):
            cell = ws.cell(row=row, column=insert_position + offset)
            if header == LOCAL_TOTAL_COL:
                cell.value = nonblank_count_formula(row, local_positions)
            else:
                cell.value = nonblank_count_formula(row, recent_score_positions.get(header, []))

        vote_vals = {column: str(ws.cell(row=row, column=column).value or "").strip() for column in vote_positions}
        row_total = sum(1 for value in vote_vals.values() if value != "")
        row_dems = sum(1 for value in vote_vals.values() if value == "D")
        row_reps = sum(1 for value in vote_vals.values() if value == "R")
        row_latest = sum(1 for column in latest_positions if vote_vals.get(column, "") != "")
        row_local = sum(1 for column in local_positions if vote_vals.get(column, "") != "")
        column_maxes["Total:"] = max(column_maxes["Total:"], row_total)
        column_maxes["Dems"] = max(column_maxes["Dems"], row_dems)
        column_maxes["REPS"] = max(column_maxes["REPS"], row_reps)
        column_maxes["Latest"] = max(column_maxes["Latest"], row_latest)
        column_maxes[LOCAL_TOTAL_COL] = max(column_maxes[LOCAL_TOTAL_COL], row_local)
        for header, positions in recent_score_positions.items():
            row_recent = sum(1 for column in positions if vote_vals.get(column, "") != "")
            column_maxes[header] = max(column_maxes[header], row_recent)

    header_to_col = {
        "Total:": insert_position, "Dems": insert_position + 1,
        "REPS": insert_position + 2, "Latest": insert_position + 3,
    }
    header_to_col.update({header: insert_position + 4 + offset for offset, header in enumerate(extra_headers)})
    build_summary_sheet(wb, ws, header_to_col, ws.max_row, column_maxes)

    if hasattr(wb, "calculation"):
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return len(latest_cols)


def score_existing_xlsx(
    input_path: Path,
    output_path: Path,
    recent_years: int,
    exclude_presidential_general: bool = True,
    score_format: str = "formulas",
) -> None:
    """Add the legacy Excel scoring columns to an already-created workbook.

    The source workbook is not modified. The four columns are inserted after
    WARD and can be written either as calculated values or Excel formulas:

      Total:  all non-blank election cells
      Dems    election cells equal to D
      REPS    election cells equal to R
      Latest  non-blank election cells dated within the recent-year window,
              excluding presidential-year (year % 4 == 0) GENERAL elections
              by default (see --include-presidential-general)
    """
    if recent_years < 0:
        raise ValueError("--recent-years must be zero or greater")
    if score_format == "values":
        latest_count = score_existing_xlsx_values(input_path, output_path, recent_years, exclude_presidential_general)
    elif score_format == "formulas":
        latest_count = score_existing_xlsx_formulas(input_path, output_path, recent_years, exclude_presidential_general)
    else:
        raise ValueError(f'Unknown score format "{score_format}"')
    print(f"Scored workbook written: {output_path}")
    print(f"Scoring columns: Total:, Dems, REPS, Latest ({latest_count} recent election columns, {score_format})")


def normalize_ward(value: object) -> str:
    """Extract the bare ward number from a WARD cell or a --ward CLI value.

    Accepts either form and reduces both to digits only, e.g. "WARREN-WARD 4"
    and "4" both normalize to "4", so the CLI value doesn't need to match the
    source file's exact "<CITY>-WARD <N>" formatting.
    """
    match = re.search(r"(\d+)", str(value or ""))
    if not match:
        raise ValueError(f'No ward number found in "{value}"')
    return match.group(1)


def filter_xlsx_by_ward(input_path: Path, ward: str, output_path: Path) -> None:
    """Filter an existing Warren voter workbook down to a single WARD.

    Works on any workbook that still has a WARD column, scored or not, so it
    can run on the raw warren-all export or on a --score-xlsx output.
    """
    df = pd.read_excel(input_path, dtype=str, keep_default_na=False)
    df.columns = [str(column).strip() for column in df.columns]
    if "WARD" not in df.columns:
        raise KeyError(f"Column WARD not found in {input_path}")

    target = normalize_ward(ward)
    result = df.loc[df["WARD"].map(normalize_ward) == target].copy()
    if result.empty:
        raise ValueError(f'No rows matched ward "{ward}" (normalized: "{target}") in {input_path}')
    for column in SCORE_VALUE_COLUMNS:
        if column in result.columns:
            converted = pd.to_numeric(result[column], errors="coerce")
            non_blank = result[column].astype(str).str.strip().ne("")
            if converted[non_blank].notna().all():
                result[column] = converted

    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(output_path, index=False, engine="openpyxl")

    from openpyxl import load_workbook as _load_workbook

    column_maxes = {
        column: int(result[column].max())
        for column in SCORE_VALUE_COLUMNS
        if column in result.columns and pd.api.types.is_numeric_dtype(result[column]) and result[column].notna().any()
    }
    header_to_col = {column: result.columns.get_loc(column) + 1 for column in column_maxes}
    wb = _load_workbook(output_path)
    build_summary_sheet(wb, wb.active, header_to_col, len(result) + 1, column_maxes)
    wb.save(output_path)

    print(f"Ward {target} workbook written: {output_path}")
    print(f"Ward {target} voters: {len(result)}")


def build_deduped_mailing_list(df: pd.DataFrame, recent_col: str) -> pd.DataFrame:
    work = df.copy()
    work["_ADDR_KEY"] = work.apply(
        lambda row: residential_address_key(combined_residential_address(row["RESIDENTIAL_ADDRESS1"], row["RESIDENTIAL_SECONDARY_ADDR"])),
        axis=1,
    )
    work = work.sort_values([recent_col, LOCAL_TOTAL_COL], ascending=[False, False], kind="stable").copy()
    reps = work.groupby("_ADDR_KEY", as_index=False, sort=False).first()

    address = reps["RESIDENTIAL_ADDRESS1"].map(fix_address)
    secondary = reps["RESIDENTIAL_SECONDARY_ADDR"].map(fix_secondary_address)
    full_address = address.where(secondary == "", address + " " + secondary)

    mailing = pd.DataFrame({
        "Recipient": reps["LAST_NAME"].str.strip().str.title() + " Household",
        "Company": [""] * len(reps),
        "Address": full_address,
        "City": reps["RESIDENTIAL_CITY"].str.strip().str.title(),
        "State": reps["RESIDENTIAL_STATE"].str.strip(),
        "Zip code": reps["RESIDENTIAL_ZIP"].str.strip(),
        # Keep the ranking signal as the final column so Vista rows can be
        # trimmed or reviewed without reopening the scored voter workbook.
        recent_col: reps[recent_col].astype(int),
    })
    return mailing.sort_values("Recipient", kind="stable")


def normalize_address(value: object) -> str:
    """Normalize an address for comparison while retaining unit information."""
    text = re.sub(r"[^A-Z0-9 ]", " ", str(value or "").upper())
    replacements = {
        r"\bNORTH\b": "N", r"\bSOUTH\b": "S", r"\bEAST\b": "E", r"\bWEST\b": "W",
        r"\bNORTHEAST\b": "NE", r"\bNORTHWEST\b": "NW", r"\bSOUTHEAST\b": "SE", r"\bSOUTHWEST\b": "SW",
        r"\bAVENUE\b": "AVE", r"\bSTREET\b": "ST", r"\bROAD\b": "RD", r"\bDRIVE\b": "DR",
        r"\bBOULEVARD\b": "BLVD", r"\bLANE\b": "LN", r"\bCOURT\b": "CT", r"\bPLACE\b": "PL",
        r"\bAPARTMENT\b": "UNIT", r"\bAPT\b": "UNIT",
    }
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text)
    return re.sub(r"\s+", " ", text).strip()


def fix_address(value: object) -> str:
    """Apply known source corrections and return a stable printable address."""
    text = normalize_address(value)
    # The SOS export has occasionally reported this Northwoods address with
    # the wrong quadrant.  Keep the correction narrow so a legitimate NW
    # address elsewhere is not changed.
    text = re.sub(r"^3820 NORTHWOODS CT NW(?=\s|$)", "3820 NORTHWOODS CT NE", text)
    # Preserve the requested correction even if it arrives in a legacy form.
    text = re.sub(r"^182 HIGH ST NW(?=\s|$)", "182 HIGH ST NE", text)
    return text


def fix_secondary_address(value: object) -> str:
    """Canonicalize apartment/unit labels for both display and deduplication."""
    return normalize_address(value)


def combined_residential_address(address: object, secondary: object) -> str:
    street = fix_address(address)
    unit = fix_secondary_address(secondary)
    return f"{street} {unit}".strip()


def load_manual_addresses(path: Path) -> pd.DataFrame:
    """Load explicitly requested addresses, if configured."""
    if not path.exists():
        return pd.DataFrame(columns=["Recipient", "Company", "Address", "City", "State", "Zip code", "VOTES_LAST_4YR"])
    manual = pd.read_csv(path, dtype=str, keep_default_na=False).fillna("")
    required = {"recipient", "address", "city", "state", "zip code"}
    missing = required - {str(column).strip().lower() for column in manual.columns}
    if missing:
        raise KeyError(f"Manual addresses CSV missing columns: {', '.join(sorted(missing))}")
    manual.columns = [str(column).strip() for column in manual.columns]
    canonical_columns = {str(column).strip().lower(): column for column in manual.columns}
    manual = manual.rename(columns={column: canonical_columns[column.lower()] for column in manual.columns})
    manual = manual.rename(columns={
        canonical_columns["recipient"]: "Recipient",
        canonical_columns["address"]: "Address",
        canonical_columns["city"]: "City",
        canonical_columns["state"]: "State",
        canonical_columns["zip code"]: "Zip code",
    })
    if "company" in canonical_columns:
        manual = manual.rename(columns={canonical_columns["company"]: "Company"})
    if "votes_last_4yr" in canonical_columns:
        manual = manual.rename(columns={canonical_columns["votes_last_4yr"]: "VOTES_LAST_4YR"})
    manual["Address"] = manual["Address"].map(fix_address)
    manual["City"] = manual["City"].map(normalize_address).str.title()
    manual["State"] = manual["State"].map(normalize_address)
    manual["Zip code"] = manual["Zip code"].str.strip()
    if "Company" not in manual:
        manual["Company"] = ""
    if "VOTES_LAST_4YR" not in manual:
        manual["VOTES_LAST_4YR"] = 0
    manual["VOTES_LAST_4YR"] = manual["VOTES_LAST_4YR"].replace("", "0").astype(int)
    return manual[["Recipient", "Company", "Address", "City", "State", "Zip code", "VOTES_LAST_4YR"]]


def add_manual_addresses(mailing: pd.DataFrame, manual: pd.DataFrame) -> pd.DataFrame:
    """Add requested addresses, replacing an existing row at the same address."""
    if manual.empty:
        return mailing
    combined = pd.concat([manual, mailing], ignore_index=True)
    keys = combined.apply(lambda row: address_key(row["Address"], row["City"], row["State"], row["Zip code"]), axis=1)
    return combined.loc[~keys.duplicated(keep="first")].copy()


def address_key(address: object, city: object, state: object, zip_code: object = "") -> str:
    """Create a city/state-aware address key; ZIP is optional for stale ZIP changes."""
    return "|".join(
        [normalize_address(address), normalize_address(city), normalize_address(state), normalize_address(zip_code)[:5]]
    )


def residential_address_key(address: object) -> str:
    """Normalize the voter-file street address used for no-delivery matching.

    Deliberately ignores city, state, and ZIP.  The source is already the
    selected voter file, and ZIP changes should not make an address exception
    stop working.
    """
    return normalize_address(address)


def verify_no_delivery_addresses(
    raw: pd.DataFrame, exceptions_path: Path, report_path: Path, clean: bool = False
) -> pd.DataFrame:
    """Audit exceptions against RESIDENTIAL_ADDRESS1 in the current voter file.

    For named official rows, also compare the exception address to the actual
    residential address of the matching voter(s).  A mismatch is only cleaned
    automatically when the official can be identified and the address is
    unambiguously different; ordinary address exceptions are never removed.
    """
    exception_keys, exceptions = load_no_delivery_addresses(exceptions_path, active_only=False)
    del exception_keys  # The audit intentionally uses street address only.
    required = {"RESIDENTIAL_ADDRESS1", "FIRST_NAME", "LAST_NAME"}
    missing = required - set(raw.columns)
    if missing:
        raise KeyError(f"Raw voter file missing columns: {', '.join(sorted(missing))}")

    voter_keys = raw["RESIDENTIAL_ADDRESS1"].map(residential_address_key)
    report_rows = []
    cleanup_indexes = []
    for index, exception in exceptions.iterrows():
        exception_address = residential_address_key(exception["address"])
        address_matches = raw.loc[voter_keys == exception_address]
        row = exception.to_dict()
        row["address_match_count"] = len(address_matches)
        row["address_match_status"] = "MATCH" if len(address_matches) else "NOT_FOUND"
        row["official_match_count"] = ""
        row["official_actual_residential_address1"] = ""
        row["official_address_status"] = "NOT_APPLICABLE"

        official = str(exception.get("official", "")).strip()
        if official:
            official_tokens = re.findall(r"[A-Z]+", official.upper())
            if len(official_tokens) >= 2:
                first_name, last_name = official_tokens[0], official_tokens[-1]
                official_matches = raw.loc[
                    raw["FIRST_NAME"].str.strip().str.upper().eq(first_name)
                    & raw["LAST_NAME"].str.strip().str.upper().eq(last_name)
                ]
                actual_addresses = list(dict.fromkeys(
                    official_matches["RESIDENTIAL_ADDRESS1"].map(str).str.strip().tolist()
                ))
                row["official_match_count"] = len(official_matches)
                row["official_actual_residential_address1"] = "; ".join(actual_addresses)
                if not len(official_matches):
                    row["official_address_status"] = "OFFICIAL_NOT_FOUND"
                elif exception_address in {
                    residential_address_key(address) for address in actual_addresses
                }:
                    row["official_address_status"] = "MATCH"
                else:
                    row["official_address_status"] = "MISMATCH"
                    is_active = str(exception.get("active", "")).strip().upper() in {"1", "TRUE", "YES", "Y"}
                    if is_active and len(actual_addresses) == 1:
                        cleanup_indexes.append(index)
        report_rows.append(row)

    report = pd.DataFrame(report_rows)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report.to_csv(report_path, index=False, quoting=csv.QUOTE_MINIMAL)

    if clean and cleanup_indexes:
        source = pd.read_csv(exceptions_path, dtype=str, keep_default_na=False)
        active_column = next(column for column in source.columns if column.strip().lower() == "active")
        for source_index in cleanup_indexes:
            source.loc[source_index, active_column] = "no"
        source.to_csv(exceptions_path, index=False, quoting=csv.QUOTE_MINIMAL)
        print(f"Deactivated {len(cleanup_indexes)} definitive official address mismatch(es): {exceptions_path}")

    print(f"Exception verification report written: {report_path}")
    return report


def load_no_delivery_addresses(path: Path, active_only: bool = True) -> tuple[set[str], pd.DataFrame]:
    """Load active address exceptions and return keys plus rows for an audit report."""
    if not path.exists():
        return set(), pd.DataFrame(columns=["address", "city", "state", "zip", "role", "official", "source", "active"])
    exceptions = pd.read_csv(path, dtype=str, keep_default_na=False).fillna("")
    required = {"address", "city", "state", "zip", "active"}
    missing = required - set(exceptions.columns.str.lower())
    if missing:
        raise KeyError(f"Exceptions CSV missing columns: {', '.join(sorted(missing))}")
    exceptions.columns = [str(column).strip().lower() for column in exceptions.columns]
    active = exceptions[exceptions["active"].str.strip().str.upper().isin({"1", "TRUE", "YES", "Y"})].copy()
    if not active_only:
        active = exceptions.copy()
    keys = {address_key(row.address, row.city, row.state, row.zip) for row in active.itertuples()}
    return keys, active


def filter_no_delivery(df: pd.DataFrame, exception_keys: set[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Remove every voter living at an exception address and return removed rows for audit."""
    keys = df.apply(
        lambda row: address_key(
            row["RESIDENTIAL_ADDRESS1"], row["RESIDENTIAL_CITY"], row["RESIDENTIAL_STATE"], row["RESIDENTIAL_ZIP"]
        ),
        axis=1,
    )
    # Exceptions are residential-address exclusions.  Match city/state too,
    # but ignore ZIP differences because voter-file ZIPs can be stale or use
    # a different ZIP for the same street address.
    exception_base_keys = {"|".join(key.split("|")[:3]) for key in exception_keys}
    match = keys.isin(exception_keys) | keys.map(lambda key: "|".join(key.split("|")[:3]) in exception_base_keys)
    return df.loc[~match].copy(), df.loc[match].copy()


def limit_mailing_list(mailing: pd.DataFrame, ranked_voters: pd.DataFrame, max_addresses: int) -> pd.DataFrame:
    """Keep the top ranked household representatives, preserving voter-score priority."""
    if max_addresses <= 0 or len(mailing) <= max_addresses:
        return mailing
    ranked = ranked_voters.drop_duplicates("_ADDR_KEY", keep="first").head(max_addresses)
    allowed = set()
    for row in ranked.itertuples():
        full_address = combined_residential_address(row.RESIDENTIAL_ADDRESS1, row.RESIDENTIAL_SECONDARY_ADDR)
        allowed.add(address_key(full_address, row.RESIDENTIAL_CITY, row.RESIDENTIAL_STATE, row.RESIDENTIAL_ZIP))
    mailing_keys = mailing.apply(lambda row: address_key(row["Address"], row["City"], row["State"], row["Zip code"]), axis=1)
    return mailing.loc[mailing_keys.isin(allowed)].copy()


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
        score_existing_xlsx(input_path, output_path, args.recent_years, args.exclude_presidential_general, args.score_format)
        return 0

    if args.ward_xlsx:
        if not args.ward:
            raise ValueError("--ward is required with --ward-xlsx")
        input_path = Path(args.ward_xlsx).expanduser().resolve()
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        if args.ward_output:
            output_path = Path(args.ward_output).expanduser().resolve()
        else:
            ward_num = normalize_ward(args.ward)
            output_path = input_path.with_name(f"{input_path.stem}-ward{ward_num}{input_path.suffix}")
        filter_xlsx_by_ward(input_path, args.ward, output_path)
        return 0

    download_dir = Path(args.download_dir).resolve()
    output_dir = Path(args.output_dir).resolve()

    raw_path = Path(args.input).expanduser().resolve() if args.input else newest_raw_file(download_dir)
    print(f"Using raw file: {raw_path}")

    df = load_raw(raw_path)
    if args.verify_exceptions:
        exception_path = Path(args.exceptions_csv).expanduser().resolve()
        report_path = output_dir / f"{datetime.today():%Y}" / f"warren-no-delivery-verification_{datetime.today():%Y-%m-%d}.csv"
        verify_no_delivery_addresses(df, exception_path, report_path, clean=args.clean_exceptions)
        return 0

    vote_cols = vote_columns(df)
    if not vote_cols:
        raise KeyError("No election columns (PRIMARY-/GENERAL-/SPECIAL-MM/DD/YYYY) found in raw file.")

    warren = df[df["CITY"].str.strip().str.upper() == args.city.upper()].copy()
    if warren.empty:
        raise ValueError(f'No rows matched CITY == "{args.city}"')
    print(f"Warren City voters: {len(warren)}")

    warren = add_scores(warren, vote_cols, args.years, args.exclude_presidential_general)
    recent_col = f"VOTES_LAST_{args.years}YR"
    warren = warren.sort_values(LOCAL_TOTAL_COL, ascending=False, kind="stable")

    if args.min_recent_votes < 0:
        raise ValueError("--min-recent-votes must be zero or greater")
    recent_voters = warren[warren[recent_col] >= args.min_recent_votes].copy()
    exception_path = Path(args.exceptions_csv).expanduser().resolve()
    exception_keys, exception_rows = load_no_delivery_addresses(exception_path)
    recent_voters, removed_voters = filter_no_delivery(recent_voters, exception_keys)
    recent_voters["_ADDR_KEY"] = recent_voters.apply(
        lambda row: residential_address_key(combined_residential_address(row["RESIDENTIAL_ADDRESS1"], row["RESIDENTIAL_SECONDARY_ADDR"])),
        axis=1,
    )
    recent_voters = recent_voters.sort_values(
        [recent_col, LOCAL_TOTAL_COL], ascending=[False, False], kind="stable"
    )
    pres_note = " (excl. presidential GENERAL)" if args.exclude_presidential_general else ""
    print(f"Voters with >={args.min_recent_votes} vote(s) in last {args.years} years{pres_note}: {len(recent_voters)}")
    print(f"Voters removed by no-delivery address exceptions: {len(removed_voters)}")

    deduped = build_deduped_mailing_list(recent_voters, recent_col)
    before_limit = len(deduped)
    deduped = limit_mailing_list(deduped, recent_voters, args.max_addresses)
    manual_path = Path(args.manual_addresses_csv).expanduser().resolve()
    deduped = add_manual_addresses(deduped, load_manual_addresses(manual_path))
    print(f"Deduped households before cap: {before_limit}")
    print(f"Mailing addresses after cap: {len(deduped)}")

    today = datetime.today()
    year_dir = output_dir / f"{today:%Y}"
    year_dir.mkdir(parents=True, exist_ok=True)
    today_str = today.strftime("%Y-%m-%d")

    all_path = year_dir / f"warren-all_{today_str}.xlsx"
    vote1_path = year_dir / f"warren-all-4yr-vote1_{today_str}.xlsx"
    vista_path = year_dir / f"warren-vista-print_{today_str}.csv"

    warren.to_excel(all_path, index=False, engine="openpyxl")
    recent_voters.to_excel(vote1_path, index=False, engine="openpyxl")
    deduped.to_csv(vista_path, index=False, quoting=csv.QUOTE_MINIMAL)
    removed_path = year_dir / f"warren-no-delivery-matches_{today_str}.xlsx"
    removed_voters.to_excel(removed_path, index=False, engine="openpyxl")
    if not exception_rows.empty:
        exception_report_path = year_dir / f"warren-no-delivery-exceptions-used_{today_str}.csv"
        exception_rows.to_csv(exception_report_path, index=False, quoting=csv.QUOTE_MINIMAL)
        print(f"Wrote: {exception_report_path}")

    print(f"Wrote: {all_path}")
    print(f"Wrote: {vote1_path}")
    print(f"Wrote: {vista_path}")
    print(f"Wrote: {removed_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
