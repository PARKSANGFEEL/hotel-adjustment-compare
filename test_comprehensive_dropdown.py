#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
New Strategy: Manipulate page size through JavaScript event handling
Since the FDS dropdown is separate from the native select,
we'll try to:
1. Trigger the dropdown opening
2. Click the 100 option programmatically
3. Monitor for actual table updates
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
        except:
            pass
    print(f"[OK] loaded {len(cookies)} cookies")
    return True

def main():
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--start-maximized')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        load_cookies(driver)
        driver.get(STATEMENTS_URL)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body')))
        time.sleep(2)

        # Try clicking statements tab
        try:
            tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-content-id="statements"]'))
            )
            tab.click()
            time.sleep(1)
        except:
            pass

        # Wait for payload
        for _ in range(20):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                break
            time.sleep(0.5)
        
        print("[INFO] Page loaded, analyzing...")
        time.sleep(2)

        # === TRY 1: Look for dropdown opening mechanism ===
        print("\n=== TRY 1: Find Dropdown Open Mechanism ===")
        result = driver.execute_script("""
            // When FDS dropdown is opened, usually a class or aria-expanded changes
            const select = document.querySelector('select.fds-field-select');
            if (!select) return { error: 'select not found' };
            
            // Check parent structure
            let parent = select.parentElement;
            const parentInfo = {
                tagName: parent.tagName,
                className: parent.className,
                attributes: Array.from(parent.attributes).map(a => ({ name: a.name, value: a.value }))
            };
            
            return { parentInfo };
        """)
        print(json.dumps(result, indent=2, ensure_ascii=False))

        # === TRY 2: Try to find and interact with actual visible dropdown ===
        print("\n=== TRY 2: Search for Open Dropdown Menu ===")
        # First, try to find any hidden dropdown that might be in a portal
        dropdown_search = driver.execute_script("""
            // Look for elements that appear like dropdown menus
            const dropdowns = document.querySelectorAll('[role="listbox"], [role="menu"], [class*="dropdown"], [class*="menu"]');
            const visibleDropdowns = Array.from(dropdowns).filter(d => {
                const style = window.getComputedStyle(d);
                return style.display !== 'none' && d.offsetParent !== null;
            });
            
            return {
                totalDropdowns: dropdowns.length,
                visibleDropdowns: visibleDropdowns.length,
                visibleDropdownDetails: visibleDropdowns.map(d => ({
                    role: d.getAttribute('role'),
                    className: d.className.substring(0, 100),
                    children: d.children.length,
                    firstChildHTML: d.innerHTML.substring(0, 200)
                })).slice(0, 5)
            };
        """)
        print(json.dumps(dropdown_search, indent=2, ensure_ascii=False))

        # === TRY 3: Check the table container for data attributes ===
        print("\n=== TRY 3: Check Table Data Attributes ===")
        table_info = driver.execute_script("""
            const table = document.querySelector('[role="table"], .fds-table');
            if (!table) return { error: 'table not found' };
            
            // Check for data-* attributes that might indicate page size
            const attributes = Array.from(table.attributes).map(a => ({
                name: a.name,
                value: a.value.substring(0, 200)
            }));
            
            // Check parent for pagination state
            const parent = table.closest('[class*="data"], [class*="list"], [class*="table"]');
            
            // Look for any element with page size info
            const pageInfo = document.querySelectorAll('[class*="page"], [class*="rows"], [class*="size"]');
            const pageElements = Array.from(pageInfo)
                .filter(p => p.offsetParent !== null)
                .map(p => ({
                    tag: p.tagName,
                    class: p.className.substring(0, 100),
                    text: p.textContent.substring(0, 50)
                }))
                .slice(0, 10);
            
            return {
                tableAttributes: attributes,
                pageElements: pageElements
            };
        """)
        print(json.dumps(table_info, indent=2, ensure_ascii=False))

        # === TRY 4: Try native select value change with all possible event triggers ===
        print("\n=== TRY 4: Comprehensive Select Value Change ===")
        change_result = driver.execute_script("""
            const select = document.querySelector('select.fds-field-select');
            if (!select) return { error: 'select not found' };
            
            const originalValue = select.value;
            
            // Set the value
            select.value = '100';
            
            // Dispatch multiple events that FDS might be listening to
            const events = ['input', 'change', 'click', 'pointerup', 'mouseup'];
            for (let eventName of events) {
                const event = new Event(eventName, { bubbles: true, cancelable: true });
                select.dispatchEvent(event);
            }
            
            // Also try with parent element
            const parent = select.parentElement;
            const parentChange = new Event('change', { bubbles: true });
            parent.dispatchEvent(parentChange);
            
            return {
                originalValue: originalValue,
                newValue: select.value,
                eventsDispatched: events
            };
        """)
        print(json.dumps(change_result, indent=2, ensure_ascii=False))

        # === TRY 5: Check if table rows changed ===
        print("\n=== TRY 5: Check Table Row Count ===")
        row_count = driver.execute_script("""
            const table = document.querySelector('[role="table"]');
            const rows = table ? table.querySelectorAll('tbody tr, [role="row"]') : [];
            return {
                rowCount: rows.length,
                firstRowText: rows[0] ? rows[0].textContent.substring(0, 100) : 'N/A'
            };
        """)
        print(json.dumps(row_count, indent=2, ensure_ascii=False))

        # === TRY 6: Try click on select element itself (as last resort) ===
        print("\n=== TRY 6: Direct Selenium Clicks ===")
        try:
            selects = driver.find_elements(By.TAG_NAME, 'select')
            if selects:
                # Try to click parent first
                select = selects[0]
                parent = select.find_element(By.XPATH, '..')
                
                try:
                    parent.click()
                    print("[OK] Clicked parent")
                except Exception as e:
                    print(f"[WARN] Parent click failed: {e}")
                
                # Try JavaScript click on select
                driver.execute_script("arguments[0].click()", select)
                print("[OK] JavaScript click on select")
                
                # Try option click
                options = select.find_elements(By.TAG_NAME, 'option')
                for opt in options:
                    if opt.get_attribute('value') == '100':
                        driver.execute_script("arguments[0].click()", opt)
                        print("[OK] JavaScript click on option 100")
                        break
                
        except Exception as e:
            print(f"[ERROR] Click attempts failed: {e}")

        print("\n[INFO] Analysis complete. Check browser window for any visual changes.")
        print("[INFO] Press Enter to close browser...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
