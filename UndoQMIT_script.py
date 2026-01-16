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
import json

# ---- Settings ----
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)
options = webdriver.ChromeOptions()
options.add_argument("--log-level=3")
options.add_argument("--silent")
service = Service(log_path="NUL")
urlUndoQMIT = "https://locadampapp01.beckman.com:8443/adamWebTier/app/undoQmit/"
LOGIN_CHECK_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/assay?assaykey=769600"
LOGIN_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/login"
temp_path = os.environ.get("TEMP") or "/tmp"
COOKIE_FILE = os.path.join(temp_path, "adam_cookies.json")
session = requests.Session()

# ---- Helpers ----
def read_assay_keys_from_txt():
    txt_path = os.path.join(temp_path, "UndoQMIT_Assays.txt")
    if not os.path.exists(txt_path):
        print(f"\033[91m❌ Error: Assay keys list not found.\033[0m")
        exit(1)
    with open(txt_path, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def cookies_are_valid(session):
    try:
        response = session.get(
            LOGIN_CHECK_URL,
            verify=False,
            timeout=15,
            allow_redirects=False
        )
        return response.status_code == 200
    except requests.exceptions.RequestException:
        return False
    
# ---- Load assays to UndoQMIT ----
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

# ---- Retrieve cookie ----
need_login = True
if os.path.exists(COOKIE_FILE):
    try:
        with open(COOKIE_FILE, "r", encoding="utf-8") as f:
            cookies = json.load(f)
        for cookie in cookies:
            session.cookies.set(cookie["name"], cookie["value"])
        if cookies_are_valid(session):
            print("\033[92m✅ You are already authenticated. Login has been skipped.\033[0m")
            need_login = False
        else:
            session.cookies.clear()
    except Exception as e:
        session.cookies.clear()

# ---- Manual login ----
if need_login:
    try:
        print("\033[92m🌐 Chrome browser will now open. Please log in to ADAM manually.\033[0m")
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_window_size(screen_width // 2, screen_height)
        driver.set_window_position(0, 0)
        driver.get(LOGIN_URL)
        print("\033[94m🔐 Waiting for login to complete...\033[0m")
        user_id_element = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.ID, "userIdData"))
        )
        print("\033[92m✅ Login successful!\033[0m")
        user_id = user_id_element.get_attribute("value")
        cookies = driver.get_cookies()
        with open(COOKIE_FILE, "w", encoding="utf-8") as f:
            json.dump(cookies, f)
        session.cookies.clear()
        for cookie in cookies:
            session.cookies.set(cookie["name"], cookie["value"])
        driver.quit()
    except TimeoutException:
        print("\033[91m❌ Error: Login timed out.\033[0m")
        driver.quit()
        exit(1)
    except WebDriverException:
        print("\033[91m❌ Chrome was closed unexpectedly.\033[0m")
        exit(1)

# ---- Process UNDO QMIT for each assay ----
print("\033[94m🔄 Starting UNDO QMIT requests...\033[0m")
total = len(assay_keys)
for index, assay_key in enumerate(assay_keys, start=1):
    print(f"\033[96m➡️  [{index}/{total}] Processing assay: {assay_key}\033[0m")
    url = urlUndoQMIT + assay_key
    try:
        response = session.get(url, verify=False, timeout=120) 
        if response.status_code == 200:
            print(f"\033[92m   ✔ UNDO QMIT successful\033[0m")
        else:
            print(f"\033[93m   ⚠ UNDO QMIT failed (HTTP {response.status_code})\033[0m")
    except requests.exceptions.RequestException as e:
        print(f"\033[91m   ❌ Request error - {e}\033[0m")        
print(f"\033[92m📊 Script completed.\033[0m")
print(f"\033[94m➡️ You may now close this window and return to the Macro Workbook.\033[0m")
time.sleep(3)
