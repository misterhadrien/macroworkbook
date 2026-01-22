# ---- Made by Hadrien Claus ----
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
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
Table1_path = os.path.join(temp_path, "Edit_ADAM_Table1.csv")
Table2_path = os.path.join(temp_path, "Edit_ADAM_Table2.csv")
Table3_path = os.path.join(temp_path, "Edit_ADAM_Table3.csv")
status_map = {
    "Pending": "PEND",
    "Confirmed": "CONF",
    "Assay Split": "SPLIT",
    "Rejected": "RJTD",
    "Cancelled": "CANC"
}
urlSaveAssay = "https://locadampapp01.beckman.com:8443/adamWebTier/app/saveAssay"
urlSaveRunOrder = "https://locadampapp01.beckman.com:8443/adamWebTier/app/saveRunorder"
headers = {"Content-Type": "application/json"}
LOGIN_CHECK_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/assay?assaykey=769600"
LOGIN_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/login"
temp_path = os.environ.get("TEMP") or "/tmp"
COOKIE_FILE = os.path.join(temp_path, "adam_cookies.json")
session = requests.Session()

# ---- Helpers ----
def safe_str(value):
    #"""Convert values safely for JSON payloads."""
    if pd.isna(value) or str(value).strip().lower() == "nan":
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
    
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
        
def get_logged_user_id(session):
    try:
        response = session.get(LOGIN_CHECK_URL, verify=False, timeout=15)
        soup = BeautifulSoup(response.text, "html.parser")
        user_id_input = soup.find("input", id="userIdData")
        if user_id_input:
            return user_id_input.get("value")
    except requests.exceptions.RequestException:
        pass
    return None
    
# ---- Header ----
print("\n" + "=" * 70)
print("📄  ADAM Editor - Macro Workbook 🧾".center(70))
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

# ---- Get User ID ----
user_id = get_logged_user_id(session)
if not user_id:
    print("\033[91m❌ Could not retrieve logged-in user ID.\033[0m")
    exit(1)

# ---- Process first table for adam assay informations ----
if os.path.exists(Table1_path):
    df = pd.read_csv(Table1_path, encoding="latin1")
    for _, row in df.iterrows():
        raw_status = safe_str(row["Status"]).strip()
        mapped_status = status_map.get(raw_status, raw_status)
        payload = {
            "assayKey": safe_str(row["Assay"]),
            "assayDesc": safe_str(row["Description"]),
            "assayComment": safe_str(row["New Comments"]),
            "assayApprovalStatus": mapped_status,
            "washBufferLotNum": safe_str(row["Wash Buffer Lot"]),
            "instrumentOperator": safe_str(row["Instrument Operator"]),
            "assayAPFRevNumber": safe_str(row["APF Rev Override"]),
            "userId": user_id
        }
        response = session.post(urlSaveAssay, headers=headers, data=json.dumps(payload), verify=False)
        if response.status_code == 200:
            print(f"✅ Assay {payload['assayKey']}: Informations updated.")
        else:
            print(f"❌ Assay {payload['assayKey']}: Error {response.status_code} - {response.text[:200]}")

# ---- Process second table for run order reagent packs ----
if os.path.exists(Table2_path):
    df = pd.read_csv(Table2_path, encoding="latin1")
    # Group by Assay to combine multiple Reagent Packs
    grouped = df.groupby("Assay")
    for assay_key, group in grouped:
        list_assay_reagent_pack = []
        for _, row in group.iterrows():
            reagent_pack = {
                "assayKey": safe_str(row["Assay"]),
                "assayReagentPckKey": safe_str(row["ReagentPack Key"]),
                "itemNum": safe_str(row["Item"]),
                "lotNum": safe_str(row["Lot"]),
                "packDesc": safe_str(row["Description"]),
                "rapidNumber": safe_str(row["Rapid"]),
                "rapidVersion": safe_str(row["Rapid Version"]),
                "pipettor": safe_str(row["Pipettor"]),
                "rowStatus": "Changed",
                "userId": user_id
            }
            list_assay_reagent_pack.append(reagent_pack)
        # Build the full payload for this assay
        payload = {
            "listAssayComponent": [],
            "listAssayReagentPack": list_assay_reagent_pack,
            "listAssayRunOrder": [],
            "deleteTOList": [],
            "overRideFlag": False
        }
        # Send POST request
        response = session.post(urlSaveRunOrder, headers=headers, data=json.dumps(payload), verify=False)
        if response.status_code == 200:
            print(f"✅ Assay {assay_key}: Reagent Packs updated.")
        else:
            print(f"❌ Assay {assay_key}: Error {response.status_code} - {response.text[:200]}")

# ---- Process third table for samples run order ----
if os.path.exists(Table3_path):
    df = pd.read_csv(Table3_path, encoding="latin1")
    # Group by Assay to combine multiple Reagent Packs
    grouped = df.groupby("Assay")
    for assay_key, group in grouped:
        list_assay_run_order = []
        for _, row in group.iterrows():
            run_order = {
                "assayKey": safe_str(row["Assay"]),
                "assayRunOrderKey": safe_str(row["Run Order Key"]),
                "sampleKey": safe_str(row["Sample Key"]),
                "runOrderSequence": safe_str(row["#"]),
                "lotNumber": safe_str(row["Lot"]),
                "repCount": safe_str(row["Reps"]),
                "calibGroup": safe_str(row["Cal Grp"]),
                "tubePosition": safe_str(row["Tube Pos"]),
                "pipettor": safe_str(row["Pipettor"]),
                "sampleComment": safe_str(row["Comment"]),
                "itemNumber": safe_str(row["Item"]),
                "sampleCategoryKey": safe_str(row["SampleCategoryKey"]),
                "rowStatus": "Changed",
                "userId": user_id
            }
            list_assay_run_order.append(run_order)
        # Build the full payload for this assay
        payload = {
            "listAssayComponent": [],
            "listAssayReagentPack": [],
            "listAssayRunOrder": list_assay_run_order,
            "deleteTOList": [],
            "overRideFlag": False
        }
        # Send POST request
        response = session.post(urlSaveRunOrder, headers=headers, data=json.dumps(payload), verify=False)
        if response.status_code == 200:
            print(f"✅ Assay {assay_key}: Samples updated.")
        else:
            print(f"❌ Assay {assay_key}: Error {response.status_code} - {response.text[:200]}")
print(f"\033[92m📊 Script completed.\033[0m")
print(f"\033[94m➡️ You may now close this window and return to the Macro Workbook.\033[0m")
time.sleep(3)
