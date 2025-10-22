# ---- Made by Hadrien Claus ----
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from datetime import datetime
import pandas as pd
import time
import os
import ctypes
import requests
import json
import urllib3

# ---- Settings ----
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)
options = webdriver.ChromeOptions()
options.add_argument("--log-level=3")
service = Service()
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
def safe_str(value):
    #"""Convert values safely for JSON payloads."""
    if pd.isna(value) or str(value).strip().lower() == "nan":
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
    
# =========================================================
# 1️ Open Chrome and wait for manual login
# =========================================================
print("\n" + "=" * 70)
print("📄  ADAM Editor - Macro Workbook 🧾".center(70))
print("=" * 70)
print("""
🔐  You will log in to ADAM manually.
🔄  After login, please wait until the script finish.
""")
print("=" * 70 + "\n")
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
# 2️ Extract cookies from Selenium to reuse with requests
# =========================================================
session = requests.Session()
for cookie in driver.get_cookies():
    session.cookies.set(cookie['name'], cookie['value'])

# =========================================================
# 3️ Process first table for adam assay informations
# =========================================================
if os.path.exists(Table1_path):
    df = pd.read_csv(Table1_path)
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

# =========================================================
# 4 Process second table for run order reagent packs
# =========================================================
if os.path.exists(Table2_path):
    df = pd.read_csv(Table2_path)
    # Group by Assay to combine multiple Reagent Packs
    grouped = df.groupby("Assay")
    for assay_key, group in grouped:
        list_assay_reagent_pack = []
        for _, row in group.iterrows():
            reagent_pack = {
                "userId": user_id,
                "assayKey": safe_str(row["Assay"]),
                "assayReagentPckKey": safe_str(row["ReagentPack Key"]),
                "itemNum": safe_str(row["Item"]),
                "lotNum": safe_str(row["Lot"]),
                "packDesc": safe_str(row["Description"]),
                "rapidNumber": safe_str(row["Rapid"]),
                "rapidVersion": safe_str(row["Rapid Version"]),
                "pipettor": safe_str(row["Pipettor"]),
                "rowStatus": "Changed"
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

# =========================================================
# 5 Process third table for samples run order
# =========================================================
if os.path.exists(Table3_path):
    df = pd.read_csv(Table3_path)
    # Group by Assay to combine multiple Reagent Packs
    grouped = df.groupby("Assay")
    for assay_key, group in grouped:
        list_assay_run_order = []
        for _, row in group.iterrows():
            run_order = {
                "userId": user_id,
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
                "rowStatus": "Changed"
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

driver.quit()
print(f"\033[92m📊 Script completed.\033[0m")
print(f"\033[94m➡️ You may now close this window and return to the Macro Workbook.\033[0m")
time.sleep(5)