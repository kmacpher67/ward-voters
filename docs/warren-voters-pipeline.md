# Warren City voter scoring & mailing list pipeline

Script: [`warren_voters_pipeline.py`](../warren_voters_pipeline.py)

Turns a raw Trumbull County SOS voter export into Warren-City-only scored lists
and a Vista Print mailing-list CSV.

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
   - `Local_Tot` — count of non-blank odd-year election columns, all-time.
     In Ohio, odd-year elections are the local/municipal cycle, so this is the
     lifetime local-vote score. This replaced the older `TOTAL_VOTES` field.
   - `VOTES_LAST_4YR` — count of non-blank election columns dated within the
     last N years (`--years`, default 4), **excluding presidential-year
     GENERAL elections** (`GENERAL-MM/DD/YYYY` where `YYYY % 4 == 0`, e.g.
     `GENERAL-11/05/2024`) by default — presidential-year turnout is not
     representative of local/municipal turnout. Pass
     `--include-presidential-general` to count them. `Local_Tot` is
     unaffected and always uses odd-year elections only.
   - Sorted by `Local_Tot` descending.
   - `→ outputs/<year>/warren-all_<date>.xlsx`
3. **Filter to local super voters**: keep rows with `VOTES_LAST_4YR >= 1` by
   default. Use `--min-recent-votes N` for a stricter definition. Remaining
   voters are ranked by recent-vote score descending, then `Local_Tot`
   descending.
   - `→ outputs/<year>/warren-all-4yr-vote1_<date>.xlsx`
4. **Apply no-delivery exceptions**: remove every voter whose residential
   address matches an active row in [`config/warren_no_delivery_addresses.csv`](../config/warren_no_delivery_addresses.csv).
   Matching normalizes case, punctuation, common street types, and directions;
   ZIP differences do not prevent an address-level match. This is
   address-level, so every registered voter at an excluded official's address
   is removed. Audit files are written for removed voters and active exceptions.
5. **Dedupe to one row per household**: group the recent-voter list by a
   canonicalized `RESIDENTIAL_ADDRESS1 + RESIDENTIAL_SECONDARY_ADDR` (trim
   whitespace, remove punctuation, standardize directions/street types, and
   treat `UNIT`, `APT`, and `APARTMENT` equivalently). Keep the highest-`Local_Tot` voter per address as the
   household representative, and emit it in the Vista mailing-list template
   format (`Vista_ListTemplate.xlsx`): `Recipient, Company, Address, City,
   State, Zip code`, with `Recipient = "<Last_Name> Household"`.
   - `→ outputs/<year>/warren-vista-print_<date>.csv`
   - The CSV uses the Vista columns `Recipient, Company, Address, City, State,
     Zip code`, followed by `VOTES_LAST_4YR` as the final review column.
   - By default, keep only the top 3,000 ranked addresses. Use
     `--max-addresses 0` for no cap.

## Usage

```bash
python3 warren_voters_pipeline.py
python3 warren_voters_pipeline.py --input "downloads/TRUMBULL (1).txt"
python3 warren_voters_pipeline.py --years 4 --output-dir outputs
python3 warren_voters_pipeline.py --min-recent-votes 2 --max-addresses 3000
python3 warren_voters_pipeline.py --max-addresses 2500
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
> 3,000-address cap. Sort remaining households by `VOTES_LAST_4YR` descending,
> then `Local_Tot` descending, dedupe by residential street address plus unit,
> and retain only the first 3,000 addresses. Export an audit file of removed
> voter IDs and the exception row that caused each removal. Do not use mailing
> address as a fallback for the official-home exclusion.

### Adjusting the number of addresses

The default run produces 3,000 deduped households. To see the cutoff and
further trim the list, change `--max-addresses`; the final
`VOTES_LAST_4YR` column in the CSV shows the recent local-vote score for each
retained household. For example:

```bash
python3 warren_voters_pipeline.py --max-addresses 2500
python3 warren_voters_pipeline.py --max-addresses 2000 --min-recent-votes 2
```

The no-delivery exceptions are applied before deduplication and before the cap.
Every active address in
[`config/warren_no_delivery_addresses.csv`](../config/warren_no_delivery_addresses.csv)
is removed from the voter rows, including all voters sharing that residential
address. The removed rows are written to
`warren-no-delivery-matches_<date>.xlsx` for review.

To score an existing all-city workbook, use the same script's XLSX mode:

```bash
python3 warren_voters_pipeline.py \
  --score-xlsx outputs/2026/warren-all_2026-09-15.xlsx \
  --score-output outputs/2026/warren-all-scored_2026-09-15.xlsx
```

This preserves the source and inserts `Total:`, `Dems`, `REPS`, `Latest`, and
`Local_Tot` after `WARD`. If the source already has a `VOTES_LAST_*YR` column,
that column is moved next to these score columns too, ahead of the election
history columns. By default the inserted columns are stored as numeric values,
not live Excel formulas:

- `Total:` counts all nonblank election columns.
- `Dems` counts election columns equal to `D`.
- `REPS` counts election columns equal to `R`.
- `Latest` counts nonblank election columns in the last six calendar years,
  excluding presidential-year GENERAL elections by default, same rule as
  `VOTES_LAST_4YR` above.
- `Local_Tot` counts nonblank odd-year election columns only. If the source
  workbook has the old `TOTAL_VOTES` column, it is replaced with this
  recalculated `Local_Tot` value.

Use `--recent-years N` to change the `Latest` window; use
`--include-presidential-general` to count presidential-year GENERAL elections
in `Latest` too. Use `--score-format formulas` to write the inserted score
columns as live Excel formulas instead:

```bash
python3 warren_voters_pipeline.py \
  --score-xlsx outputs/2026/warren-all_2026-09-15.xlsx \
  --score-output outputs/2026/warren-all-scored-formulas_2026-09-15.xlsx \
  --score-format formulas
```

Formula mode writes `=COUNTA(...)`, `=COUNTIF(...,"D")`,
`=COUNTIF(...,"R")`, and a row-specific `Latest` sum of nonblank recent
election cells. It also writes formula versions of `Local_Tot` and any
`VOTES_LAST_*YR` column present in the source workbook. Open the workbook in
Excel or LibreOffice to calculate and display formula results. Value mode is
safer for viewers that do not recalculate formulas. The ward-filter mode
preserves these score columns as numbers instead of rewriting them as text.

To pull a single ward out of an existing workbook (raw or scored), use the
same script's ward-filter mode:

```bash
python3 warren_voters_pipeline.py \
  --ward-xlsx outputs/2026/warren-all-scored_2026-09-15.xlsx \
  --ward 4 \
  --ward-output outputs/2026/warren-ward4-scored_2026-09-15.xlsx
```

This preserves the source workbook and writes only the rows whose `WARD`
column matches. Matching is on the ward number only (`normalize_ward()`
extracts digits from both the `--ward` value and each `WARD` cell), so `4`
and `WARREN-WARD 4` are equivalent. Omitting `--ward-output` defaults to
`<input-stem>-ward<N>.xlsx`.

## Notes / open questions

- Household dedup key is address-only (last-name-at-household is not part of
  the key), matching the request to dedupe "based on RESIDENTIAL_ADDRESS1,
  RESIDENTIAL_SECONDARY_ADDR". Two unrelated voters sharing a duplex address
  with identical secondary-address text will collapse into one household.
- Known manual corrections are applied before deduplication: Pennock's address
  is `182 HIGH ST NE`, the Northwoods correction is `3820 NORTHWOODS CT NE`,
  and printed unit labels use `UNIT`. Explicit additions are read from
  `config/warren_manual_addresses.csv` and retained after the cap.
- This pipeline is independent of the ward-specific scoring script
  (`voters-warren-scored.py`), which adds Excel-formula-driven Total/Dems/
  REPS/Muni/Latest columns per ward file instead of a pandas-computed score
  across the whole city. Use `voters-warren-scored.py` when a party (D/R)
  breakdown or per-ward Excel workbook with live formulas is needed instead.
