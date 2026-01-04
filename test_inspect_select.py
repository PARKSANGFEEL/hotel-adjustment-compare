# -*- coding: utf-8 -*-
"""
Test: inspect select element and its wrapper for actual clickable UI
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
        except Exception as e:
            print(f"[WARN] add_cookie failed {c.get('name')}: {e}")
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

        # click statements tab
        try:
            tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-content-id="statements"]'))
            )
            tab.click()
            time.sleep(1)
        except Exception:
            pass

        for _ in range(20):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                break
            time.sleep(0.5)

        # find first select with options 10,25,50,100
        selects = driver.find_elements(By.TAG_NAME, 'select')
        select = None
        for s in selects:
            opts = [o.get_attribute('value') for o in s.find_elements(By.TAG_NAME, 'option')]
            if '100' in opts:
                select = s
                break
        
        if not select:
            print("[ERROR] select not found")
            return

        print("[OK] select found")
        print(f"[INFO] select display: {driver.execute_script('return arguments[0].offsetParent === null', select)}")
        print(f"[INFO] select class: {select.get_attribute('class')}")
        print(f"[INFO] select style: {select.get_attribute('style')}")
        
        # check parent/wrapper
        parent_html = driver.execute_script("""
            const sel = arguments[0];
            const p = sel.parentElement;
            return {
                tag: p.tagName,
                class: p.className,
                visible: p.offsetParent !== null,
                innerHTML_snippet: p.innerHTML.substring(0, 200)
            };
        """, select)
        print(f"[INFO] parent element: {parent_html}")
        
        # look for visible button/div near select
        siblings = driver.execute_script("""
            const sel = arguments[0];
            const parent = sel.parentElement;
            const result = [];
            for (let el of parent.querySelectorAll('button, div[role="button"], [role="combobox"]')) {
                if (el.offsetParent !== null) {
                    result.push({
                        tag: el.tagName,
                        class: el.className,
                        text: el.innerText.substring(0,50),
                        visible: true,
                        xpath: el.getAttribute('aria-label')
                    });
                    if (result.length >= 3) break;
                }
            }
            return result;
        """, select)
        print(f"[INFO] visible siblings: {siblings}")
        
        # Try clicking the visible button
        if siblings:
            try:
                visible_button = driver.execute_script("""
                    const sel = arguments[0];
                    const parent = sel.parentElement;
                    for (let el of parent.querySelectorAll('button, div[role="button"], [role="combobox"]')) {
                        if (el.offsetParent !== null) {
                            el.click();
                            return true;
                        }
                    }
                    return false;
                """, select)
                print(f"[INFO] clicked visible element: {visible_button}")
                time.sleep(0.5)
            except Exception as e:
                print(f"[ERROR] click failed: {e}")

    finally:
        input("\n[INFO] Press Enter to close browser...")
        driver.quit()

if __name__ == '__main__':
    main()
