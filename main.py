import pandas as pd
import time
import os 

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Credentials and paths
EXCEL_PATH = "PlaceHolder.xlsx"
SISENSE_URL = "https://your-sisense-url" 
USERNAME = "user_placeholder" 
PASSWORD = "pw_placeholder"  
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")  
MACRO_SECTOR = "PlaceHolder"

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Configure Edge WebDriver
options = webdriver.EdgeOptions()
options.use_chromium = True 
preferences = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "profile.default_content_settings.popups": 0,
    "directory_upgrade": True
}
options.add_experimental_option("prefs", preferences)
driver = webdriver.Edge(options=options)
wait = WebDriverWait(driver, 20)

# Step 1: Log into Sisense
driver.get(SISENSE_URL)

wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(USERNAME)
driver.find_element(By.NAME, "password").send_keys(PASSWORD + Keys.RETURN)

# Step 2: Find the folder
wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Advisory Views')]"))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Ownership Dashboards')]"))).click() 
wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Institutional Trends')]"))).click() 

# Step 3: Read Excel file
df = pd.read_excel(EXCEL_PATH)

# Step 4: Apply filters 
holding_filter = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(text(),'Holding Source')]/following-sibling::div[contains(@class,'left-arrow')]")))
holding_filter.click()
time.sleep(1)  
research = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Research')]")))
research.click()

# Step 5: Apply firm filters
for firm in df["Firm Name"]:
    print(f"Processing firm: {firm}")
    firm_filter = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(text(),'Firm Name')]/following-sibling::div[contains(@class,'left-arrow')]")))
    firm_filter.click()
    time.sleep(2)
    firm_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Start typing to search...']")))
    firm_input.clear()
    firm_input.send_keys(firm)
    time.sleep(2)
    firm_confirm = wait.until(EC.presence_of_all_elements_located((By.XPATH, f"//div[contains(text(), '{firm}')]")))
    if firm_confirm >= 2:
        firm_confirm[1].click()
    else:
        firm_confirm[0].click()
    time.sleep(2) 

    macro_filter = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(text(),'Macro Sector')]/following-sibling::div[contains(@class,'left-arrow')]")))
    macro_filter.click()
    time.sleep(2)
    macro_sector = wait.until(EC.presence_of_element_located((By.XPATH, f"//div[contains(text(), '{MACRO_SECTOR}')]")))
    macro_sector.click()
    time.sleep(2)
    
    # Step 6: Download the file
    find_chart = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Macro Sector Movements')]")))
    chart_area = find_chart.find_element(By.XPATH, "//ancestor::div[contains(@class, 'chart-container') or contains(@class, 'chart-area') or contains(@class, 'widget-container')]")  # Uncertain element class name
    three_dots = chart_area.find_element(By.XPATH, ".//button[contains(@class, 'menu-button') or contains(@aria-label, 'More options')]")  # Uncertain element class name
    three_dots.click()
    time.sleep(2)
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Download')]")))
    download_button.click()
    time.sleep(1)
    download_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'CSV File')]")))
    time.sleep(10)
    
    print(f"Download complete for firm: {firm}")
    
print("All files processed")
driver.quit()