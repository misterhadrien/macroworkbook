# ---- Made by Hadrien Claus ----
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
import time
import os
import ctypes
import requests
import urllib3

# ---- Settings ----
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)
options = webdriver.ChromeOptions()
options.add_argument("--log-level=3")
service = Service()
urlUndoQMIT = "https://locadampapp01.beckman.com:8443/adamWebTier/app/undoQmit/"

def read_assay_keys_from_txt():
    temp_path = os.environ.get("TEMP") or "/tmp"
    txt_path = os.path.join(temp_path, "UndoQMIT_Assays.txt")
    if not os.path.exists(txt_path):
        print(f"\033[91m❌ Error: Assay keys list not found.\033[0m")
        exit(1)
    with open(txt_path, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]
    
# =========================================================
# Open Chrome and wait for manual login
# =========================================================
print("\n" + "=" * 70)
print("📄  ADAM Undo QMIT - Macro Workbook 🧾".center(70))
print("=" * 70)
print("""
🔐  You will log in to ADAM manually.
🔄  After login, please wait until the script finish.
""")
print("=" * 70 + "\n")
assay_keys = read_assay_keys_from_txt()
print(f"\033[92m📚 Loaded {len(assay_keys)} ADAM assay{'s' if len(assay_keys) != 1 else ''}.\033[0m")
try:
    print("\033[92m🌐 Chrome browser will now open. Please log in to ADAM manually.\033[0m")
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(screen_width // 2, screen_height)
    driver.set_window_position(0, 0)
    driver.get("https://locadampapp01.beckman.com:8443/adamWebTier/login")
    print("\033[94m🔐 Waiting for login to complete...\033[0m")
    try:
        user_id_element = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.ID, "userIdData"))
        )
    except TimeoutException:
        print("\033[91m❌ Error: Login timed out.\033[0m")
        driver.quit()
        exit(1)
    print("\033[92m✅ Login successful!\033[0m")
    user_id = user_id_element.get_attribute("value")
except WebDriverException as e:
    time.sleep(1)
    print("\n\033[91m❌ Error: Chrome browser was closed unexpectedly.\033[0m")
    time.sleep(5)
except Exception as e:
    time.sleep(1)
    print(f"\033[91m❌ Error: An unexpected error occurred: {e}\033[0m")
    time.sleep(5)

# =========================================================
# Extract cookies from Selenium to reuse with requests
# =========================================================
session = requests.Session()
for cookie in driver.get_cookies():
    session.cookies.set(cookie['name'], cookie['value'])

# =========================================================
# Process undo qmit for each assay
# =========================================================
print("\033[94m🔄 Starting Undo QMIT requests...\033[0m")
for assay_key in assay_keys:
    url = urlUndoQMIT + assay_key
    try:
        response = session.get(url, verify=False, timeout=120) 
        if response.status_code == 200:
            print(f"\033[92m✅ Undo QMIT successful for assay: {assay_key}\033[0m")
        else:
            print(f"\033[93m⚠️ Undo QMIT failed for assay: {assay_key} (HTTP {response.status_code})\033[0m")
    except requests.exceptions.RequestException as e:
        print(f"\033[91m❌ Undo QMIT request error for assay: {assay_key} - {e}\033[0m")
        
driver.quit()
print(f"\033[92m📊 Script completed.\033[0m")
print(f"\033[94m➡️ You may now close this window and return to the Macro Workbook.\033[0m")
time.sleep(5)
