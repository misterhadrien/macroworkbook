# ---- Made by Hadrien Claus ----
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import time
import os
import ctypes
import requests
import urllib3
import json
import msvcrt

# ---- Settings ----
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)
options = webdriver.ChromeOptions()
options.add_argument("--log-level=3")
options.add_argument("--silent")
service = Service(log_path="NUL")
ASSAY_REPORTS_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/assayReportPage/"
PASS_REPORT_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/passFailIndReport"
APPROVE_REPORT_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/approveIndReport"
LOGIN_CHECK_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/app/assay?assaykey=769600"
LOGIN_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/login"
temp_path = os.environ.get("TEMP") or "/tmp"
COOKIE_FILE = os.path.join(temp_path, "adam_cookies.json")
session = requests.Session()

# ---- Helpers ----
def read_assay_report_pairs():
    txt_path = os.path.join(temp_path, "Report_Keys.txt")
    if not os.path.exists(txt_path):
        print("\033[91m❌ Error: Report keys list not found.\033[0m")
        exit(1)
    pairs = []
    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            if "," not in line:
                continue
            assay, report = line.strip().split(",", 1)
            pairs.append((assay.strip(), int(report.strip())))
    return pairs
        
def get_report_timestamp(html, target_report_key):
    soup = BeautifulSoup(html, "html.parser")
    row = soup.find("tr", attrs={"data-reportkey": str(target_report_key)})
    if not row:
        return None, None, None
    PassFailStatus = row.get("data-passfailcode")
    ApprovalStatus = row.get("data-approvalstatus")
    timestamp = row.get("data-reporttimestamp")
    return PassFailStatus, ApprovalStatus, timestamp

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
    
def input_password(prompt="Password: "):
    print(prompt, end="", flush=True)
    password = ""
    while True:
        ch = msvcrt.getch()
        if ch in {b'\r', b'\n'}:  # Enter pressed
            print()
            break
        elif ch == b'\x08':  # Backspace
            if len(password) > 0:
                password = password[:-1]
                print("\b \b", end="", flush=True)
        else:
            password += ch.decode("utf-8")
            print("*", end="", flush=True)
    return password

# ---- Load assays ----
print("\n" + "=" * 70)
print("📄  ADAM Pass & Approve Reports - Macro Workbook 🧾".center(70))
print("=" * 70)
print("""
🔐  You will log in to ADAM manually.
➡️  After login, please enter your password for report approval.
""")
print("=" * 70 + "\n")
pairs = read_assay_report_pairs()
total = len(pairs)
print(f"\033[92m📚 Loaded {total} assay report{'s' if total != 1 else ''}.\033[0m")
    
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

# ---- Get User ID & password for approval requests ----
approver_user_id = get_logged_user_id(session)
if not approver_user_id:
    print("\033[91m❌ Could not retrieve logged-in user ID.\033[0m")
    exit(1)
print("""
🔐 Your ADAM password is required to approve reports.
➡️ Enter your password to approve reports, or press Enter to skip approval requests.
""")
approver_password = input_password(f"Enter password for {approver_user_id} (press Enter to skip): ").strip()
if not approver_password:
    print("\033[93m⚠ Approval requests will be skipped.\033[0m")
    approver_password = None

# ---- Process pass & approve report for each assay ----
print("\033[94m🔄 Starting Pass & Approve report requests...\033[0m")
for index, (assay_key, report_key) in enumerate(pairs, start=1):
    print(f"\033[96m➡️  [{index}/{total}] Processing Assay {assay_key} (Report Key {report_key})\033[0m")
    time.sleep(1)
    url = ASSAY_REPORTS_URL + assay_key
    try:
        response = session.get(url, verify=False, timeout=120)
        if response.status_code != 200:
            print(f"\033[91m   ❌ Failed to load assay report page")
            continue
        PassFailStatus, ApprovalStatus, timestamp = get_report_timestamp(response.text, report_key)
        if not timestamp:
            print(f"\033[91m   ❌ Report key not found on assay report page")
            continue
        # ---- Pass report if Pending ----
        if PassFailStatus == "Pending":
            payload = {"passFailCode": "PASS", "passFailReasonCode": "", "passFailComment": "", "reportKeyList": [report_key], "resultsTimeStamp": [timestamp]}
            pass_response = session.post(PASS_REPORT_URL, json=payload, verify=False, timeout=30)
            if pass_response.status_code == 200:
                print(f"\033[92m   ✔ Assay report set to PASS\033[0m")
                if approver_password and ApprovalStatus == "Pending":
                    time.sleep(1)
                    response = session.get(url, verify=False, timeout=120)
                    if response.status_code != 200:
                        print(f"\033[91m   ⚠  Approval failed: Error when loading assay report page")
                        continue
                    PassFailStatus, ApprovalStatus, timestamp = get_report_timestamp(response.text, report_key)
                    if not timestamp:
                        print(f"\033[91m   ⚠  Approval failed: Report key not found")
                        continue
            else:
                print(f"\033[93m   ⚠ Failed to set to PASS (HTTP {pass_response.status_code})\033[0m")
        else:
            print(f"   ℹ  Report Already Pass")
        # ---- Approve report if conditions met ----
        if approver_password:
            if PassFailStatus != "Pending" and ApprovalStatus == "Pending":
                approve_payload = {
                    "passFailCode": "CONF",
                    "passFailReasonCode": "",
                    "passFailComment": "",
                    "approverUserID": approver_user_id,
                    "approverPassword": approver_password,
                    "reportKeyList": [report_key],
                    "resultsTimeStamp": [timestamp],
                    "assayKey": assay_key
                }
                approve_response = session.post(APPROVE_REPORT_URL, json=approve_payload, verify=False, timeout=30)
                if approve_response.status_code == 200:
                    print(f"\033[92m   ✔ Assay report approved\033[0m")
                else:
                    response_text = approve_response.text.strip()
                    if response_text == "Invalid User credential!":
                        print(f"\033[93m   ⚠  Approval Failed: Incorrect password. Approval requests will be skipped.\033[0m")
                        approver_password = None
                    else:
                        print(f"\033[93m   ⚠  Approval failed (HTTP {approve_response.status_code}: {response_text})\033[0m")
            else:
                if ApprovalStatus == "Confirmed":
                    print(f"   ℹ  Report Already Approved")
                else:
                    print(f"   ℹ  Report not eligible for approval (PassFail: {PassFailStatus}, Approval: {ApprovalStatus})")
    except requests.exceptions.RequestException as e:
        print(f"\033[91m   ❌ Request error - {e}\033[0m")
print(f"\033[92m📊 Script completed.\033[0m")
print(f"\033[94m➡️ You may now close this window and return to the Macro Workbook.\033[0m")
print(f"\033[94m➡️ Click on the Search button to refresh assays report list.\033[0m")
time.sleep(5)