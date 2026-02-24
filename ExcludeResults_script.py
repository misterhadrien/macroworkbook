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
import json
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
temp_path = os.environ.get("TEMP") or "/tmp"
json_path = os.path.join(temp_path, "ExcludeResults_payload.json")
headers = {"Content-Type": "application/json"}
LOGIN_CHECK_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/assay?assaykey=769600"
LOGIN_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/login"
EXCLUDE_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/excludeResults"
temp_path = os.environ.get("TEMP") or "/tmp"
COOKIE_FILE = os.path.join(temp_path, "adam_cookies.json")
session = requests.Session()

# ---- Helpers ----
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
    
# ---- Header ----
print("\n" + "=" * 70)
print("📄  ADAM Exclude Results - Macro Workbook 🧾".center(70))
print("=" * 70)
print("""
🔐  You will log in to ADAM manually.
🔄  After login, please wait until the script finish.
""")
print("=" * 70 + "\n")

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
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.ID, "userIdData"))
        )
        print("\033[92m✅ Login successful!\033[0m")
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

# ---- Process exclude results ----
if os.path.exists(json_path):
    response = session.get(LOGIN_CHECK_URL, verify=False, timeout=120)
    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)  # Load JSON as Python dict
    exclude = session.post(EXCLUDE_URL, headers=headers, json=payload, verify=False)
    if exclude.status_code == 200:
        print(f"✅ Results excluded successfully.")
    else:
        print(f"❌ Error {exclude.status_code} - {exclude.text[:200]}")
print(f"\033[92m📊 Script completed.\033[0m")
print(f"\033[94m➡️ You may now close this window and return to the Macro Workbook.\033[0m")
time.sleep(3)