from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Set up Chrome options
options = Options()
options.headless = True  # Run in headless mode (no GUI)
# You can add more options to mimic a regular browser if necessary

# Initialize the driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Navigate to the URL
# url = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:HOME::::::"
url = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:DOWNLOAD::FILE:NO:2:P2_PRODUCT_NUMBER:78"
driver.get(url)

# You can now extract page content or interact with the page as needed
page_content = driver.page_source
print(page_content)

driver.quit()
