# Warren City voter scoring & mailing list pipeline

Script: [`warren_voters_pipeline.py`](../warren_voters_pipeline.py)

Turns a raw Trumbull County SOS voter export into three Warren-City-only outputs:
scored full list, recent-voter subset, and a deduped household mailing list.

## Input

Raw file: comma-delimited `TRUMBULL*.txt` from the Ohio SOS voter file download
(product 78), same format used by `download_trumbull_ward.py` /
`voters_warren.py`. Columns include one per election
(`PRIMARY-MM/DD/YYYY`, `GENERAL-MM/DD/YYYY`, `SPECIAL-MM/DD/YYYY`), with the
voter's party code (`D`/`R`/etc.) in that column if they voted in it.

The Ohio SOS site returns HTTP 403 for plain `requests` downloads (see
`voters_warren.py`), and this sandbox could not complete a live Selenium
download via `download_trumbull_ward.py` either (no working browser network
egress here). By default the pipeline reuses the newest `TRUMBULL*.txt`
already in `downloads/`. Pass `--input <path>` to pin a specific file, or run
`download_trumbull_ward.py` first if a live re-download is needed on a
machine where SOS access works.

## What it does

1. **Load & filter**: read the raw `.txt`, keep rows where `CITY == "WARREN CITY"`
   (all 7 wards, not just one).
2. **Score every voter**:
   - `TOTAL_VOTES` — count of non-blank election columns, all-time.
   - `VOTES_LAST_4YR` — count of non-blank election columns dated within the
     last N years (`--years`, default 4), **excluding presidential-year
     GENERAL elections** (`GENERAL-MM/DD/YYYY` where `YYYY % 4 == 0`, e.g.
     `GENERAL-11/05/2024`) by default — presidential-year turnout is not
     representative of local/municipal turnout. Pass
     `--include-presidential-general` to count them. `TOTAL_VOTES` is
     unaffected and always includes presidential-year elections.
   - Sorted by `TOTAL_VOTES` descending.
   - `→ outputs/<year>/warren-all_<date>.xlsx`
3. **Filter to local super voters**: keep rows with `VOTES_LAST_4YR >= 1` by
   default. Use `--min-recent-votes N` for a stricter definition. Remaining
   voters are ranked by recent-vote score descending, then `TOTAL_VOTES`
   descending.
   - `→ outputs/<year>/warren-all-4yr-vote1_<date>.xlsx`
4. **Apply no-delivery exceptions**: remove every voter whose residential
   address matches an active row in [`config/warren_no_delivery_addresses.csv`](../config/warren_no_delivery_addresses.csv).
   Matching normalizes case, punctuation, common street types, and directions;
   a blank ZIP in an exception matches any ZIP at that address. This is
   address-level, so every registered voter at an excluded official's address
   is removed. Audit files are written for removed voters and active exceptions.
5. **Dedupe to one row per household**: group the recent-voter list by
   `RESIDENTIAL_ADDRESS1 + RESIDENTIAL_SECONDARY_ADDR` (street address + unit/apt,
   exact match after trim/uppercase — `UNIT 3` and `APT 3` at the same street
   number are treated as different households since the SOS data doesn't
   normalize those). Keep the highest-`TOTAL_VOTES` voter per address as the
   household representative, and emit it in the Vista mailing-list template
   format (`Vista_ListTemplate.xlsx`): `Recipient, Company, Address, City,
   State, Zip code`, with `Recipient = "<Last_Name> Household"`.
   - `→ outputs/<year>/warren-all-4yr-vote1-deduped_<date>.xlsx`
   - By default, keep only the top 2,000 ranked addresses. Use
     `--max-addresses 0` for no cap.

## Usage

```bash
python3 warren_voters_pipeline.py
python3 warren_voters_pipeline.py --input "downloads/TRUMBULL (1).txt"
python3 warren_voters_pipeline.py --years 4 --output-dir outputs
python3 warren_voters_pipeline.py --min-recent-votes 2 --max-addresses 2000
python3 warren_voters_pipeline.py --exceptions-csv config/warren_no_delivery_addresses.csv
python3 warren_voters_pipeline.py --input downloads/TRUMBULL.txt --verify-exceptions
python3 warren_voters_pipeline.py --input downloads/TRUMBULL.txt --verify-exceptions --clean-exceptions
```

`--verify-exceptions` audits every exception row against the selected raw
voter file using normalized `RESIDENTIAL_ADDRESS1` only; city, state, and ZIP
are not used to decide whether the address exists. Named elected-official rows
are additionally checked against the official's actual residential address in
that voter file. The report is written to
`outputs/<year>/warren-no-delivery-verification_<date>.csv`.
`--clean-exceptions` deactivates only definitive named-official mismatches;
ordinary address exceptions and ambiguous official matches are left unchanged.

The CSV is editable. Add rows with `address,city,state,zip,role,official,source,active`;
use `active=no` to pause a row without deleting its history.

Recommended copy/paste upgrade instruction:

> Before mailing, load `config/warren_no_delivery_addresses.csv`. Normalize both
> voter residential addresses and exception addresses (uppercase, remove
> punctuation, standardize street types and directions), then remove every
> voter whose normalized address/city/state/ZIP matches an active exception.
> Perform this before household deduplication and before applying the
> 2,000-address cap. Sort remaining households by `VOTES_LAST_4YR` descending,
> then `TOTAL_VOTES` descending, dedupe by residential street address plus unit,
> and retain only the first 2,000 addresses. Export an audit file of removed
> voter IDs and the exception row that caused each removal. Do not use mailing
> address as a fallback for the official-home exclusion.

To score an existing all-city workbook, use the same script's XLSX mode:

```bash
python3 warren_voters_pipeline.py \
  --score-xlsx outputs/2026/warren-all_2026-09-15.xlsx \
  --score-output outputs/2026/warren-all-scored_2026-09-15.xlsx
```

This preserves the source and inserts `Total:`, `Dems`, `REPS`, and `Latest`
after `WARD`. The columns are calculated values: all-time nonblank votes,
`D` votes, `R` votes, and nonblank votes in the last six calendar years
(excluding presidential-year GENERAL elections by default, same rule as
`VOTES_LAST_4YR` above), respectively. Change the recent window with
`--recent-years N`; use `--include-presidential-general` to count
presidential-year GENERAL elections in `Latest` too.

## Notes / open questions

- Household dedup key is address-only (last-name-at-household is not part of
  the key), matching the request to dedupe "based on RESIDENTIAL_ADDRESS1,
  RESIDENTIAL_SECONDARY_ADDR". Two unrelated voters sharing a duplex address
  with identical secondary-address text will collapse into one household.
- This pipeline is independent of the ward-specific scoring script
  (`voters-warren-scored.py`), which adds Excel-formula-driven Total/Dems/
  REPS/Muni/Latest columns per ward file instead of a pandas-computed score
  across the whole city. Use `voters-warren-scored.py` when a party (D/R)
  breakdown or per-ward Excel workbook with live formulas is needed instead.
