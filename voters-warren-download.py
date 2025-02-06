import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ----------------------------------
# 1. Setup Download Directory and Chrome Options
# ----------------------------------
download_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

chrome_options = Options()
prefs = {
    "download.default_directory": download_dir,  # Set the download folder
    "download.prompt_for_download": False,         # Download without prompting
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)
# Uncomment the next line to run in headless mode (if needed)
# chrome_options.headless = True

# ----------------------------------
# 2. Use Selenium to Download the File
# ----------------------------------
# Initialize the Selenium WebDriver with Chrome
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

# URL for TRUMBULL county file (county #78) that returns a 403 via requests
download_url = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:DOWNLOAD::FILE:NO:2:P2_PRODUCT_NUMBER:78"
driver.get(download_url)

# ----------------------------------
# 3. Wait for the Download to Complete
# ----------------------------------
# Poll the download folder for a .txt file (adjust the extension if needed)
timeout = 60  # seconds
start_time = time.time()
downloaded_file_path = None

while time.time() - start_time < timeout:
    files = os.listdir(download_dir)
    txt_files = [f for f in files if f.lower().endswith('.txt')]
    if txt_files:
        # If more than one file exists, choose the first one
        downloaded_file_path = os.path.join(download_dir, txt_files[0])
        # (Optionally, you can check for temporary download file names or extensions)
        break
    time.sleep(1)

if not downloaded_file_path:
    print("File download timed out. Exiting.")
    driver.quit()
    exit()

print(f"Downloaded file: {downloaded_file_path}")

# Close the Selenium browser
driver.quit()

# ----------------------------------
# 4. Process the Downloaded File with Pandas
# ----------------------------------
# Adjust the delimiter if your file is not comma-separated
try:
    df = pd.read_csv(downloaded_file_path, delimiter=",")
except Exception as e:
    print(f"Error reading the file with pandas: {e}")
    exit()

# Filter rows where the "CITY" column contains "WARREN CITY" (case-insensitive)
filtered_df = df[df["CITY"].str.contains("WARREN CITY", case=False, na=False)]

# Sort the filtered data by the "PRECINCT" column
sorted_df = filtered_df.sort_values(by="PRECINCT_NAME")

# ----------------------------------
# 5. Save the Processed Data to an Excel File
# ----------------------------------
today_str = datetime.today().strftime("%Y-%m-%d")
# output_filename = f"CityOfWarren{today_str}.xls"
output_filename = f"CityOfWarren{today_str}.xlsx"
sorted_df.to_excel(output_filename, index=False, engine='openpyxl')


try:
    # Using the xlwt engine to produce an .xls file
    # sorted_df.to_excel(output_filename, index=False, engine='xlwt')
    print(f"Filtered data saved to {output_filename}")
except Exception as e:
    print(f"Error saving the Excel file: {e}")
