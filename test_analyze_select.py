# -*- coding: utf-8 -*-
"""
Analyze the select element and page framework.
Inspect actual event listeners and JavaScript state.
"""
import json
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

BASE_DIR = Path(__file__).resolve().parent
COOKIES_FILE = BASE_DIR / "expedia_cookies.json"
STATEMENTS_URL = "https://apps.expediapartnercentral.com/lodging/accounting/statementsAndInvoices.html?htid=17300293&tab=invoices"

def load_cookies(driver):
    if not COOKIES_FILE.exists():
        print("[ERROR] cookies file missing")
        return False
    cookies = json.loads(COOKIES_FILE.read_text())
    driver.get("https://www.expediapartnercentral.com/Account/Logon?signedOff=true")
    time.sleep(2)
    for c in cookies:
        try:
            driver.add_cookie(c)
        except Exception:
            pass
    print(f"[OK] loaded {len(cookies)} cookies")
    return True

def main():
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    try:
        load_cookies(driver)
        driver.get(STATEMENTS_URL)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body')))

        # 명세서 탭 클릭
        try:
            tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-content-id="statements"]'))
            )
            tab.click()
            time.sleep(2)
            print("[OK] 명세서 탭 clicked")
        except Exception as e:
            print(f"[WARN] tab click failed: {e}")

        # wait payload
        for _ in range(20):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                break
            time.sleep(0.5)
        print("[INFO] payload ready")

        # find select
        selects = driver.find_elements(By.TAG_NAME, 'select')
        print(f"[INFO] found {len(selects)} select elements")
        
        select_elem = None
        for i, select in enumerate(selects):
            opts = [o.get_attribute('value') for o in select.find_elements(By.TAG_NAME, 'option')]
            print(f"  [{i}] options: {opts}")
            if '100' in opts:
                select_elem = select
                print(f"       ^ This is the pagination select")

        if not select_elem:
            print("[ERROR] no select with option 100 found")
            return

        # Inspect the select element
        print("\n=== SELECT ELEMENT ANALYSIS ===")
        print(f"Tag: {select_elem.tag_name}")
        print(f"ID: {select_elem.get_attribute('id')}")
        print(f"Class: {select_elem.get_attribute('class')}")
        print(f"Name: {select_elem.get_attribute('name')}")
        print(f"Current value: {select_elem.get_attribute('value')}")
        
        # Check for data attributes
        all_attrs = driver.execute_script("return Object.keys(arguments[0].attributes).map(i => arguments[0].attributes[i].name)", select_elem)
        print(f"All attributes: {all_attrs}")
        
        # Check parent structure
        parent_html = driver.execute_script("return arguments[0].parentElement.outerHTML.substring(0, 500)", select_elem)
        print(f"Parent HTML (first 500 chars): {parent_html}...")
        
        # Detect framework
        print("\n=== FRAMEWORK DETECTION ===")
        frameworks = {
            'React': "!!window.__REACT_DEVTOOLS_GLOBAL_HOOK__",
            'Vue': "!!window.__VUE__",
            'Angular': "!!window.ng",
            'jQuery': "!!window.jQuery",
        }
        for fw, check in frameworks.items():
            result = driver.execute_script(f"return {check}")
            print(f"{fw}: {result}")
        
        # Try to get event listeners
        print("\n=== EVENT LISTENERS ===")
        listeners = driver.execute_script("""
            const select = arguments[0];
            const listeners = getEventListeners(select);
            return Object.keys(listeners || {});
        """, select_elem)
        print(f"Event listeners: {listeners if listeners else 'Unable to get (getEventListeners not available)'}")
        
        # Check if it's a React-controlled input
        print("\n=== REACT/VUE CHECK ===")
        is_react_controlled = driver.execute_script("""
            const select = arguments[0];
            // Check for React fiber
            const fiberKey = Object.keys(select).find(key => key.startsWith('__react'));
            if (fiberKey) {
                const fiber = select[fiberKey];
                console.log('React fiber found:', fiber);
                return 'React-controlled: YES';
            }
            // Check for Vue
            if (select.__vue__) {
                return 'Vue-controlled: YES';
            }
            return 'Not controlled by framework (or data hidden)';
        """, select_elem)
        print(f"{is_react_controlled}")
        
        # Try clicking on the actual select element and checking DOM changes
        print("\n=== TESTING SELECT INTERACTION ===")
        print("Initial rows in table:")
        rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
        print(f"  {len(rows)} rows")
        
        print("\nManually clicking select and waiting 5 seconds for interaction...")
        driver.execute_script("arguments[0].scrollIntoView();", select_elem)
        time.sleep(0.5)
        select_elem.click()
        print("  [clicked select]")
        
        time.sleep(5)
        print("\nAfter 5 second wait, you should manually select '100' from dropdown.")
        print("Press ENTER once you've selected it or timeout expires...")
        
        try:
            input()
        except:
            pass
        
        # Check final state
        print("\nFinal state after manual interaction:")
        print(f"Select value: {select_elem.get_attribute('value')}")
        rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
        print(f"Row count: {len(rows)}")
        
        print("\n=== KEEPING BROWSER OPEN ===")
        print("Browser will stay open. Press ENTER to close...")
        try:
            input()
        except:
            pass

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
