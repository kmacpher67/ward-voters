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
     last N years (`--years`, default 4).
   - Sorted by `TOTAL_VOTES` descending.
   - `→ outputs/<year>/warren-all_<date>.xlsx`
3. **Filter to recent voters**: keep rows with `VOTES_LAST_4YR >= 1`.
   - `→ outputs/<year>/warren-all-4yr-vote1_<date>.xlsx`
4. **Dedupe to one row per household**: group the recent-voter list by
   `RESIDENTIAL_ADDRESS1 + RESIDENTIAL_SECONDARY_ADDR` (street address + unit/apt,
   exact match after trim/uppercase — `UNIT 3` and `APT 3` at the same street
   number are treated as different households since the SOS data doesn't
   normalize those). Keep the highest-`TOTAL_VOTES` voter per address as the
   household representative, and emit it in the Vista mailing-list template
   format (`Vista_ListTemplate.xlsx`): `Recipient, Company, Address, City,
   State, Zip code`, with `Recipient = "<Last_Name> Household"`.
   - `→ outputs/<year>/warren-all-4yr-vote1-deduped_<date>.xlsx`

## Usage

```bash
python3 warren_voters_pipeline.py
python3 warren_voters_pipeline.py --input "downloads/TRUMBULL (1).txt"
python3 warren_voters_pipeline.py --years 4 --output-dir outputs
```

To score an existing all-city workbook, use the same script's XLSX mode:

```bash
python3 warren_voters_pipeline.py \
  --score-xlsx outputs/2026/warren-all_2026-09-15.xlsx \
  --score-output outputs/2026/warren-all-scored_2026-09-15.xlsx
```

This preserves the source and inserts `Total:`, `Dems`, `REPS`, and `Latest`
after `WARD`. The columns are calculated values: all-time nonblank votes,
`D` votes, `R` votes, and nonblank votes in the last six calendar years,
respectively. Change the recent window with `--recent-years N`.

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
