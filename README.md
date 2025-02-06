# ward-voters
Ohio Sos secretary of state county voter database filter for warren city, score voter's activity and save to ward base files. 

wrote several programs in stages to figure out the python code.

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

