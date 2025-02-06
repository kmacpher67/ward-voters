import os
import time
from datetime import datetime
import pandas as pd

# --- Selenium and WebDriver Imports ---
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- openpyxl for postprocessing the Excel file ---
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

#########################################
# STEP 1. Download the file via Selenium
#########################################

# Set up the download directory
download_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

# Configure Chrome options to automatically download files.
chrome_options = Options()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)
# You may run without headless mode to ensure downloads work reliably:
# chrome_options.headless = True

# Initialize the Selenium Chrome WebDriver.
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

# URL that returns the TRUMBULL county file (which otherwise gives a 403 error with requests)
download_url = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:DOWNLOAD::FILE:NO:2:P2_PRODUCT_NUMBER:78"
driver.get(download_url)

# Wait for the file download (poll the download_dir for a .txt file)
timeout = 60  # seconds
start_time = time.time()
downloaded_file_path = None

while time.time() - start_time < timeout:
    files = os.listdir(download_dir)
    txt_files = [f for f in files if f.lower().endswith('.txt')]
    if txt_files:
        downloaded_file_path = os.path.join(download_dir, txt_files[0])
        break
    time.sleep(1)

if not downloaded_file_path:
    print("File download timed out. Exiting.")
    driver.quit()
    exit()

print(f"Downloaded file: {downloaded_file_path}")
driver.quit()

######################################################
# STEP 2. Process the downloaded file with pandas
######################################################
try:
    # Adjust delimiter if necessary (here we assume CSV format)
    df = pd.read_csv(downloaded_file_path, delimiter=",")
except Exception as e:
    print(f"Error reading the file with pandas: {e}")
    exit()

# (Optional) Print original columns for debugging
# print("Original columns:", df.columns.tolist())

# Normalize column names: remove extra whitespace and convert to uppercase.
df.columns = df.columns.str.strip().str.upper()
# For example, "Ward" becomes "WARD" and "Primary-03/07/2000" becomes "PRIMARY-03/07/2000".

# Filter rows to include only those with "WARREN CITY" in the CITY column (case-insensitive).
# filtered_df = df[df["CITY"].str.contains("WARREN CITY", case=False, na=False)]
# instead of filtering the raw county data by City, in the future have it filter by "WARD" containing the "WARREN-WARD"
filtered_df = df[df["WARD"].str.contains("WARREN-WARD", case=False, na=False)]

# Sort by the PRECINCT column. (Make sure PRECINCT exists in your data.)
try:
    sorted_df = filtered_df.sort_values(by="PRECINCT_NAME")
except KeyError:
    print("The column 'PRECINCT' was not found. Check the file headers.")
    print("Available columns:", filtered_df.columns.tolist())
    exit()

#############################################
# STEP 3. Write the DataFrame to an Excel file
#############################################
today_str = datetime.today().strftime("%Y-%m-%d")
output_filename = f"CityOfWarren{today_str}.xlsx"

# Write to Excel using openpyxl engine (so we can later modify formulas)
try:
    sorted_df.to_excel(output_filename, index=False, engine='openpyxl')
    print(f"Data written to {output_filename}")
except Exception as e:
    print(f"Error saving the Excel file: {e}")
    exit()

#########################################################################
# STEP 4. Insert 4 columns (Total:, Dems, REPS, Muni) and add formulas.
#########################################################################
# Open the workbook with openpyxl
wb = load_workbook(output_filename)
ws = wb.active

# --- Identify key columns based on header row (assumed to be row 1) ---

# Find the column index for "WARD" and for the first vote column "PRIMARY-03/07/2000".
# (Assuming exact header text; adjust if needed.)
ward_col_idx = None
primary_col_idx = None

for col in range(1, ws.max_column + 1):
    cell_val = ws.cell(row=1, column=col).value
    if isinstance(cell_val, str):
        cell_val = cell_val.strip().upper()
        if cell_val == "WARD":
            ward_col_idx = col
        if cell_val == "PRIMARY-03/07/2000":
            primary_col_idx = col

if ward_col_idx is None:
    print("Could not find 'WARD' column in the header. Exiting.")
    exit()

if primary_col_idx is None:
    print("Could not find 'PRIMARY-03/07/2000' column in the header. Exiting.")
    exit()

# For our insertion, we want to insert four new columns immediately to the right of the WARD column.
insert_position = ward_col_idx + 1
ws.insert_cols(insert_position, amount=4)

# Because we inserted columns before the vote columns, the vote columns have shifted to the right.
# Adjust the primary vote column index:
new_primary_idx = primary_col_idx + 4

# Identify the last column in the worksheet (assumed to be the last vote column).
last_vote_col_idx = ws.max_column

# Determine which columns (by index) in the vote range correspond to an odd year.
# We examine the header row (row 1) for columns from new_primary_idx to last_vote_col_idx.
odd_year_cols = []
for col in range(new_primary_idx, last_vote_col_idx + 1):
    header_val = ws.cell(row=1, column=col).value
    if isinstance(header_val, str) and header_val.upper().startswith("PRIMARY-"):
        # Attempt to extract the year from the last 4 characters of the header.
        try:
            year = int(header_val.strip()[-4:])
            if year % 2 == 1:
                odd_year_cols.append(col)
        except:
            pass

# Determine the Excel column letters for the vote range.
first_vote_letter = get_column_letter(new_primary_idx)
last_vote_letter = get_column_letter(last_vote_col_idx)

# Also get the letters for the odd-year vote columns.
odd_year_letters = [get_column_letter(c) for c in odd_year_cols]

# --- Write the new header labels in the inserted columns ---
new_headers = ["Total:", "Dems", "REPS", "Muni"]
for i, header in enumerate(new_headers):
    ws.cell(row=1, column=insert_position + i, value=header)

# --- For each data row, insert formulas in the new columns ---
for row in range(2, ws.max_row + 1):
    # Build the range reference for vote columns (e.g., "F2:K2")
    vote_range = f"${first_vote_letter}${row}:${last_vote_letter}${row}"

    # Column for Total: count nonblank cells in the vote range.
    total_formula = f"=COUNTA({vote_range})"
    ws.cell(row=row, column=insert_position, value=total_formula)

    # Column for Dems: count cells equal to "D".
    dems_formula = f'=COUNTIF({vote_range},"D")'
    ws.cell(row=row, column=insert_position + 1, value=dems_formula)

    # Column for REPS: count cells equal to "R".
    reps_formula = f'=COUNTIF({vote_range},"R")'
    ws.cell(row=row, column=insert_position + 2, value=reps_formula)

    # Column for Muni: sum IFs over each odd-year vote column.
    # We build a formula like: =IF(F2="D",1,0)+IF(H2="D",1,0)+...
    muni_parts = []
    for col_letter in odd_year_letters:
        muni_parts.append(f'IF({col_letter}{row}="D",1,0)')
    muni_formula = "=" + "+".join(muni_parts) if muni_parts else "=0"
    ws.cell(row=row, column=insert_position + 3, value=muni_formula)

# Save the workbook with the formulas inserted.
wb.save(output_filename)
print(f"Post-processing complete. Final file saved as {output_filename}")
