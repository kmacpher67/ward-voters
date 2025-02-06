import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Define the directory where you want to save the downloaded file.
# Here we create a folder named 'downloads' in the current working directory.
download_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

# Set up Chrome options to automatically download files to the specified directory.
chrome_options = Options()
prefs = {
    "download.default_directory": download_dir,  # Set default download directory
    "download.prompt_for_download": False,         # Do not prompt for download
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True                   # Enable safe browsing
}
chrome_options.add_experimental_option("prefs", prefs)

# Note: Running in headless mode might complicate file downloads.
# It's recommended to run with the browser visible for file downloads.
# chrome_options.headless = True  # Uncomment if you need headless mode, but test carefully.

# Initialize the Chrome WebDriver.
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# URL for the TRUMBULL download (county #78)
url = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:DOWNLOAD::FILE:NO:2:P2_PRODUCT_NUMBER:78"

# Navigate to the URL to trigger the file download.
driver.get(url)

# Wait for a few seconds to allow the download to complete.
# You might want to increase the wait time for larger files or implement a loop that checks for file completion.
time.sleep(10)

# Optionally, check if the file exists in the download directory.
downloaded_files = os.listdir(download_dir)
print("Downloaded files:", downloaded_files)

# Clean up: close the browser.
driver.quit()
