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

=IF(CC2="D",1,0)+IF(CL2="D",1,0)+IF(CW2="D",1,0)+IF(DD2="D",1,0)+IF(DJ2="D",1,0)+IF(DO2="D",1,0)+IF(DW2="D",1,0)+IF(CC2="D",1,0)+IF(CD2="D",1,0)+IF(CJ2="D",1,0)+IF(CK2="D",1,0)+IF(CO2="D",1,0)+IF(CP2="D",1,0)+IF(CQ2="D",1,0)+IF(CU2="D",1,0)+IF(CV2="D",1,0)+IF(DB2="D",1,0)+IF(DC2="D",1,0)+IF(DH2="D",1,0)+IF(CW2="D",1,0)+IF(DM2="D",1,0)+IF(DD2="D",1,0)+IF(DO2="D",1,0)+IF(DJ2="D",1,0)+IF(DW2="D",1,0)


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

