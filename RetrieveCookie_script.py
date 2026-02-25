# ---- Made by Hadrien Claus ----
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
import os
import json
import ctypes

# ---- Settings ----
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)
options = webdriver.ChromeOptions()
options.add_argument("--log-level=3")
options.add_argument("--silent")
service = Service(log_path="NUL")
temp_path = os.environ.get("TEMP") or "/tmp"
LOGIN_URL = "https://locadampapp01.beckman.com:8443/adamWebTier/login"
temp_path = os.environ.get("TEMP") or "/tmp"
COOKIE_FILE = os.path.join(temp_path, "adam_cookies.json")

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
    driver.quit()
except TimeoutException:
    print("\033[91m❌ Error: Login timed out.\033[0m")
    driver.quit()
except WebDriverException:
    print("\033[91m❌ Chrome was closed unexpectedly.\033[0m")