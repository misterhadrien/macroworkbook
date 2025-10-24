# ---- Made by Hadrien Claus ----
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from datetime import datetime
import pandas as pd
import time
import os
import ctypes

# ---- Settings ----
EXCEL_FILENAME = "audit_trail_export.xlsx"
TABLE_SELECTOR = "div#audittrail_0_auditreport_0 table.contentBorder"
ROWS_SELECTOR = "table.contentBorder tr.contentBackground"
NEXT_BUTTON_NAME = "audittrail_0_pager1_next_0"
column_names = ["Document", "Title", "Date", "Time Zone", "Version", "User", "Event"]
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)

# ---- Setup ----
options = webdriver.ChromeOptions()
options.add_argument("--log-level=3")
service = Service()

def read_document_names_from_txt():
    temp_path = os.environ.get("TEMP") or "/tmp"
    txt_path = os.path.join(temp_path, "EDMS_documents_names.txt")
    if not os.path.exists(txt_path):
        print(f"\033[91m❌ Error: Document names not found.\033[0m")
        exit(1)
    with open(txt_path, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def search_document_by_name(document_name, current_index, total_documents, js_code):
    driver.switch_to.default_content()
    driver.switch_to.frame(2)
    driver.switch_to.frame(2)
    driver.switch_to.frame(2)
    try:
        old_row = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "Search60_doclistgrid_0_0"))
        )
    except:
        old_row = None
        
    driver.switch_to.default_content()
    driver.switch_to.frame(1)
    title_text = None
    try:
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtSearch"))
        )
        search_box.clear()
        search_box.send_keys(document_name)
        search_box.send_keys(Keys.ENTER)
        print(f"\033[94m🔍 Searching for document {current_index + 1} of {total_documents}: {document_name}\033[0m")
        driver.switch_to.default_content()
        driver.switch_to.frame(2)
        driver.switch_to.frame(2)
        driver.switch_to.frame(2)
        if old_row:
            WebDriverWait(driver, 10).until(EC.staleness_of(old_row))
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "Search60_doclistgrid_0_0")),
            )
        except TimeoutException:
            no_data_elements = driver.find_elements(By.ID, "searchnodata")
            if no_data_elements:
                print(f"\033[91m❌ No documents found for '{document_name}'.\033[0m")
                return None
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.selectable")
        exact_matches = []
        partial_matches = []
        for row in rows:
            try:
                tds = row.find_elements(By.TAG_NAME, "td")
                exact_match_found = False
                partial_match_found = False
                for i in [2, 4]:  # Only check td[2] and td[4]
                    try:
                        # Try to get full text from span[title]
                        span = tds[i].find_element(By.CSS_SELECTOR, "span[title]")
                        text = span.get_attribute("title").strip()
                    except:
                        # Fallback to visible text
                        text = tds[i].text.strip()
                    if not exact_match_found and text == document_name:
                        exact_match_found = True
                    elif not partial_match_found and document_name.lower() in text.lower():
                        partial_match_found = True
                if exact_match_found:
                    exact_matches.append(row)
                elif partial_match_found:
                    partial_matches.append(row)
            except Exception as e:
                print(f"\033[91m❌ Error while searching for {document_name}: {e}\033[0m")
                return None
        if len(exact_matches) == 1:
            print(f"\033[92m✅ Found and opening Audit Trail of {document_name}\033[0m")
            exact_matches[0].click()
        else:
            if len(exact_matches) > 1:
                print(f"\033[93m⚠️ Multiple exact matches for '{document_name}'\033[0m")
            elif partial_matches:
                print(f"\033[93m⚠️ Partial matches found for '{document_name}'\033[0m")
            else:
                print(f"\033[91m❌ No match found for '{document_name}'\033[0m")
            
            print("\033[93m➡️ Please manually select the correct document row...\033[0m")
            try:
                WebDriverWait(driver, 120).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "tr.selectable.selected"))
                )
            except TimeoutException:
                print("\033[91m❌ Timeout: No row was selected.\033[0m")
                return None
        selected_row = driver.find_element(By.CSS_SELECTOR, "tr.selectable.selected")
        tds = selected_row.find_elements(By.TAG_NAME, "td")
        try:
            title_span = tds[4].find_element(By.CSS_SELECTOR, "span.dmfStrLenFrmtr > span[title]")
            title_text = title_span.get_attribute("title")
        except:
            title_text = tds[4].text.strip()
        driver.execute_script(js_code)
        return title_text
    except TimeoutException:
        print(f"\033[91m⏳ Timeout while searching for '{document_name}'\033[0m")
        return None
    except Exception as e:
        print(f"\033[91m❌ Error while searching for {document_name}: {e}\033[0m")
        return None

def extract_all_pages(title):
    audit_data = []
    page = 1
    while True:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(2)
            driver.switch_to.frame(2)
            driver.switch_to.frame(2)
            driver.switch_to.frame(0)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, TABLE_SELECTOR))
            )
            rows = driver.find_elements(By.CSS_SELECTOR, ROWS_SELECTOR)
            doc_name_element = driver.find_element(By.CLASS_NAME, "dialogFileName")
            document_name = doc_name_element.text.strip()
            page_data = []
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = [cell.text.strip() for cell in cells]
                row_data.insert(0, document_name)
                row_data.insert(1, title if title is not None else "")
                page_data.append(row_data)
            print(f"\033[92m✅ {len(page_data)} rows extracted from Audit Trail Page {page}.\033[0m")
            audit_data.extend(page_data)
            page += 1
            try:
                next_button = driver.find_element(By.NAME, NEXT_BUTTON_NAME)
                old_table = driver.find_element(By.CSS_SELECTOR, TABLE_SELECTOR)
                driver.execute_script("arguments[0].click();", next_button)
                WebDriverWait(driver, 10).until(EC.staleness_of(old_table))
            except NoSuchElementException:
                print("\033[92m➡️ Audit Trail extraction complete.\033[0m")
                break
        except TimeoutException:
            print("\033[91m❌ Timeout waiting for audit trail table.\033[0m")
            break
        except Exception as e:
            print(f"\033[91m❌ Error: {e}\033[0m")
            break
    return audit_data

# ---- Main Execution ----
print("\n" + "=" * 70)
print("📄  EDMS Audit Trail Extractor - Macro Workbook 🧾".center(70))
print("=" * 70)
print("""
🔐  You will log in to EDMS manually.
🔄  All Audit Trail pages for each documents are processed automatically.
📌  If no exact match is found during document search,
    you must manually select the correct document in the list.
""")
print("=" * 70 + "\n")
doc_names = read_document_names_from_txt()
print(f"\033[92m📚 Loaded {len(doc_names)} document{'s' if len(doc_names) != 1 else ''}.\033[0m")
try:
    print("\033[92m🌐 Chrome browser will now open. Please log in to EDMS manually.\033[0m")
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(screen_width // 2, screen_height)
    driver.set_window_position(0, 0)
    driver.get("https://edms.beckman.com/edms/component/main")
    print("\033[94m🔐 Waiting for login to complete...\033[0m")
    try:
        WebDriverWait(driver, 300).until(
            EC.invisibility_of_element_located((By.NAME, "Login_Button_0"))
        )
    except TimeoutException:
        print("\033[91m❌ Error: Login timed out.\033[0m")
        driver.quit()
        exit(1)
    print("\033[92m✅ Login successful!\033[0m")
    try:
        # Wait until frame 1 is available and switch to it
        WebDriverWait(driver, 10).until(lambda d: len(d.find_elements(By.TAG_NAME, "frame")) > 1)
        driver.switch_to.frame(1)
        # Wait until search box is present
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtSearch"))
        )
    except TimeoutException:
        print("\033[91m❌ Error: Timed out waiting for page to load after login.\033[0m")
        driver.quit()
        exit(1)
    all_data = []
    processed = 0
    # Get Audit Trail button ID
    driver.switch_to.default_content()
    driver.switch_to.frame(2)
    driver.switch_to.frame(2)
    driver.switch_to.frame(1)
    menu_bar_form = driver.find_element(By.ID, "MenuBar_0")
    client_id_full = menu_bar_form.find_element(By.NAME, "__dmfRequestId").get_attribute("value")
    client_id_prefix = client_id_full.split("~~")[0]
    audit_trail_id = client_id_prefix + "_MenuBar_doc_audittrail_0"
    js_code = f'fireDynamicActionEvent("{audit_trail_id}")'
    for i, doc_name in enumerate(doc_names):
        try:
            title = search_document_by_name(doc_name, i, len(doc_names), js_code)
            if title is None:
                continue
            data = extract_all_pages(title)
            if data:
                processed += 1
                all_data.extend(data)
                ok_button = driver.find_element(By.NAME, "ComboContainer_cancel_0")
                ok_button.click()
            else:
                print("\033[91m⚠️ No rows found in Audit Trail.\033[0m")
        except Exception as e:
            print(f"\033[91m❌ Error processing {doc_name}: {e}\033[0m")
    driver.quit()
    if all_data:
        df = pd.DataFrame(all_data, columns=column_names)
        out_path = os.path.join(os.environ.get("TEMP") or "/tmp", EXCEL_FILENAME)
        df.to_excel(out_path, index=False)
        print(f"\033[92m📊 Export completed successfully.\033[0m")
        print(f"\033[94m➡️ You may now close this window and return to the Macro Workbook.\033[0m")
        print(f"\033[94m💡 Then click on 'Get Versions' to continue processing.\033[0m")
    else:
        print("\033[91m❌ Error: No data to export.\033[0m")
except WebDriverException as e:
    print("\n\033[91m❌ Error: Chrome browser was closed unexpectedly.\033[0m")
except Exception as e:
    print(f"\033[91m❌ Error: An unexpected error occurred: {e}\033[0m")
time.sleep(5)
