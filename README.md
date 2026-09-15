# ward-voters
Ohio Sos secretary of state county voter database filter for warren city, score voter's activity and save to ward base files. 

wrote several programs in stages to figure out the python code.

Found a problem with the formulas, they are hard coded the column calculations $BA### sorting breaks badly. 




## main code to run 
this program runs selenium and downloads the .txt file and converts them to xlsx file filtering just the city of warren
then calculates the scores of tot, D, R, muni voters. 
```
python voters_warren-scored.py 
```

## Trumbull ward download
Use this helper when you want the latest Trumbull County SOS voter file and a Warren ward filter in one step.

Default: Warren City Ward 4
```
bash download_trumbull_ward.sh
```

Other ward example:
```
bash download_trumbull_ward.sh --ward 3
```

If the SOS site blocks automated access, you can filter a local raw file instead:
```
bash download_trumbull_ward.sh --input 'downloads/TRUMBULL (1).txt' --ward 4
```

Outputs are written under `outputs/<year>/` so each run stays grouped by year.

## Warren City scored + deduped mailing list

See [docs/warren-voters-pipeline.md](docs/warren-voters-pipeline.md) for full
details. Filters the raw Trumbull SOS file to Warren City (all wards), scores
every voter by total lifetime votes and votes in the last 4 years, and
produces a household-deduped mailing list in the Vista template format.

```
python3 warren_voters_pipeline.py
```

The Vista output is a CSV at `outputs/<year>/warren-vista-print_<date>.csv`
(not an `.xls`/`.xlsx` file). It contains up to 3,000 households, deduped on
`RESIDENTIAL_ADDRESS1` plus `RESIDENTIAL_SECONDARY_ADDR`, with active addresses
from `config/warren_no_delivery_addresses.csv` removed before the cap. The
final `VOTES_LAST_4YR` column shows the local-vote score used to rank the list.
Use `--max-addresses N` to finesse the count, or `--max-addresses 0` for all
matching households:

```
python3 warren_voters_pipeline.py --max-addresses 2500
python3 warren_voters_pipeline.py --max-addresses 2000 --min-recent-votes 2
```

To add the legacy party/activity score columns to an existing all-city workbook
without changing the input file:

```
python3 warren_voters_pipeline.py \
  --score-xlsx outputs/2026/warren-all_2026-09-15.xlsx \
  --score-output outputs/2026/warren-all-scored_2026-09-15.xlsx
```

This inserts `Total:`, `Dems`, `REPS`, and `Latest` immediately after `WARD`.
`Latest` counts nonblank election columns from the last six calendar years;
use `--recent-years N` to change that window. The source workbook is preserved.

## Verify no-delivery addresses

Before mailing, verify the no-delivery list against the latest raw voter file:

```
python3 warren_voters_pipeline.py \
  --input downloads/TRUMBULL.txt \
  --verify-exceptions
```

Verification compares normalized `RESIDENTIAL_ADDRESS1` values only; city,
state, and ZIP are not used to decide whether an address is present. Named
officials are also checked against their actual residential address in the
voter file. The audit is written under `outputs/<year>/`.

To deactivate stale, definitive elected-official address rows while preserving
the original rows for audit history, add `--clean-exceptions`:

```
python3 warren_voters_pipeline.py \
  --input downloads/TRUMBULL.txt \
  --verify-exceptions \
  --clean-exceptions
```

The normal pipeline automatically uses only rows whose `active` value is
`yes`. Re-run verification whenever a newer SOS voter file is downloaded.

Notes from kens google drive on how to integate google maps: 
https://docs.google.com/document/d/1Miosc88rydmc6TaZL_I1nJBV-QTW3dCpGULZ1GwGor4/edit?tab=t.0

See the scratch sheet for the various notes and LLM command snippets used to generate the code. 


## basic logic from chatgpt 

weaking column insertions

I'm adjusting the script to add 6 columns instead of 4 between "WARD" and "VOTE". This involves shifting the primary vote column index and updating new column headers.

Adjusting labels

Aligning the column headers, I’m adding "Latest" and "Both" as the 5th and 6th columns. The new_headers list will now have 6 items, ensuring the correct sequence.

Working on formulas

I’m mapping out formulas for the updated columns: "Total:" uses a vote count, "Dems" and "REPS" use counts for specific values, "Muni" sums odd-year votes, and "Latest" summarizes votes for the past 6 years.

Tracking recent votes

I’m focusing on counting columns with embedded dates from within the past 6 years to gauge recent votes. This involves evaluating columns with labels like "PRIMARY-05/07/2019" to "SPECIAL-01/07/2025".

Setting thresholds

I’m thinking the threshold should be the current year minus 6. Columns with headers from "PRIMARY-05/07/2019" to "SPECIAL-01/07/2025" will have their votes counted or summed.

Weighing options

I’m considering whether "calculates any votes" means counting non-blank cells or applying a COUNTIF formula to cells with dates above the threshold. This determines the method of counting.

Revisiting calculations

I'm re-reading the instructions to clarify that the "Latest" column will sum votes from columns with dates in the last 6 years, aligning with labels from "PRIMARY-05/07/2019" to "SPECIAL-01/07/2025".

Mapping out voting columns

To calculate votes for the "Latest" column, I'll count non-blank cells in the contiguous range from new_primary_idx to last_vote_col_idx for columns with dates above the threshold.



Muni columns: 

2023 May, 2021, 2019 may 

DW = PRIMARY-05/02/2023
DO = PRIMARY-05/04/2021
DJ = PRIMARY-05/07/2019
DD = PRIMARY-05/02/2017
CW = PRIMARY-05/05/2015
CL = PRIMARY-05/03/2011
CC = PRIMARY-05/05/2009

=countif(dw2,"D")+countif(DO2,"D")++countif(DJ2,"D")+countif(DD2,"D")++countif(CW2,"D")


=IF(CC2="D",1,0)+IF(CL2="D",1,0)+IF(CW2="D",1,0)+IF(DD2="D",1,0)+IF(DJ2="D",1,0)+IF(DO2="D",1,0)+IF(DW2="D",1,0)+IF(CC2="D",1,0)+IF(CD2="D",1,0)+IF(CJ2="D",1,0)+IF(CK2="D",1,0)+IF(CO2="D",1,0)+IF(CP2="D",1,0)+IF(CQ2="D",1,0)+IF(CU2="D",1,0)+IF(CV2="D",1,0)+IF(DB2="D",1,0)+IF(DC2="D",1,0)+IF(DH2="D",1,0)+IF(CW2="D",1,0)+IF(DM2="D",1,0)+IF(DD2="D",1,0)+IF(DO2="D",1,0)+IF(DJ2="D",1,0)+IF(DW2="D",1,0)


MAILER 
Valid (3449)


=IF(DN2="D",1,0)+IF(DF2="D",1,0)+IF(DA2="D",1,0)+IF(CU2="D",1,0)

LATEST: 
=COUNTIF($EA2:$DJ2, "<>")
=IF(DH2<>"",1,0)+IF(DI2<>"",1,0)+IF(DK2<>"",1,0)+IF(DM2<>"",1,0)+IF(DN2<>"",1,0)+IF(DO2<>"",1,0)+IF(DQ2<>"",1,0)+IF(DR2<>"",1,0)+IF(DT2<>"",1,0)+IF(DU2<>"",1,0)+IF(DV2<>"",1,0)+IF(DW2<>"",1,0)+IF(DX2<>"",1,0)+IF(DZ2<>"",1,0)+IF(EA2<>"",1,0)+IF(EB2<>"",1,0)+IF(EE2<>"",1,0)

DISPLAY

=CONCATENATE(D3," ",LEFT(I3,4),"T=",AW3,"D=",AX3,"R=",AY3,"M=",AZ3,"L=",BA3,"B=",BB3)
WISWELL 1951T=48D=17R=0M=7L=12B=0

StreetName
BRADFORD ST NW


Add a program for parsing filtering the wards for google maps 
I manually created the "CityOfWarren2025-02-06-target-googlemaps.csv" deleting all the unneed rows.
Use the file "CityOfWarren2025-02-06-target-googlemaps.csv" and write a python program to read this csv and  save individual CSV files by ward with a maximum size of 2000 rows per ward file, each file would be named with CityOfWarren2025-02-06-target-googlemaps-WARD1-Rows1-2000.csv where WARD1= WARD column and -Rows1-2000 would be required for Rows2001-4000 etc. Call the python program  wardfilterforgooglemaps.py 

Below is an example Python script named wardfilterforgooglemaps.py that reads the CSV file, groups the data by the "WARD" column, and then splits each ward’s data into multiple files (with up to 2000 rows per file). Each file is named following the pattern:

Format:
CityOfWarren2025-02-06-target-googlemaps-{WARD}-Rows{start}-{end}.csv
