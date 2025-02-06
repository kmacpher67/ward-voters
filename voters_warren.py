import pandas as pd
import requests
from io import StringIO
from datetime import datetime

# URL of the data file from the Ohio Secretary of State website
url = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:DOWNLOAD::FILE:NO:2:P2_PRODUCT_NUMBER:78"
headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "en-US,en;q=0.9,ja;q=0.8"
}

url1 = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:HOME::::::"
response = requests.get(url1, headers=headers)
response.raise_for_status()  # This will raise an HTTPError for bad responses (e.g., 403)

# Download the file. Note: the file format (delimiter, encoding, etc.) may require adjustments.
response = requests.get(url, headers=headers)
response.raise_for_status()  # Raise an error if the download failed

# Assume the file is a text file with delimited data and a header row.
# You may need to change the delimiter depending on the actual file structure.
# For example, if the file is comma-separated, use delimiter=','.
# If it is tab-separated, use delimiter='\t'.
data = StringIO(response.text)
df = pd.read_csv(data, delimiter=",")  # Adjust delimiter if needed

# Filter rows where the CITY column contains "WARREN CITY"
# If you need an exact match, use == instead of str.contains.
filtered_df = df[df["CITY"].str.contains("WARREN CITY", na=False)]

# Sort the filtered DataFrame by the PRECINCT column.
# If PRECINCT is numeric but stored as a string, you might need to convert it.
sorted_df = filtered_df.sort_values(by="PRECINCT")

# Create the filename with today’s date in YYYY-MM-DD format.
today_str = datetime.today().strftime("%Y-%m-%d")
filename = f"CityOfWarren{today_str}.xls"

# Save the DataFrame to an Excel file.
# Note: Pandas will actually create an XLSX file if using the default Excel writer.
# To force the .xls extension, you might need to use a library like xlwt (which supports only older Excel formats)
# Here we use the .xls extension as requested, but note that the file may be in XLSX format.
sorted_df.to_excel(filename, index=False)

print(f"Filtered data saved to {filename}")
