#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Strategy: Find and interact with FDS framework dropdown.
The native select is hidden, so we need to:
1. Find the actual visible dropdown trigger button
2. Click it to open the dropdown
3. Select the "100" option from the visible dropdown list
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
from selenium.webdriver.common.action_chains import ActionChains
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

        # Click statements tab if needed
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
        
        print("[INFO] Page loaded")
        time.sleep(2)

        # === STRATEGY 1: Look for the actual dropdown button by analyzing FDS structure ===
        print("\n=== STRATEGY 1: Find FDS Dropdown Button ===")
        result = driver.execute_script("""
            // FDS dropdown structure: usually has a hidden select + visible button
            // The button might be a sibling or a wrapper
            
            const select = document.querySelector('select.fds-field-select');
            if (!select) return { error: 'select not found' };
            
            // Method A: Look for button with data-target or aria-controls pointing to this select
            const selectId = select.id || '';
            console.log('Select ID:', selectId);
            
            // Method B: Look for parent container that might have a button
            let container = select.parentElement;
            for (let i = 0; i < 5; i++) {
                if (!container) break;
                
                // Check for visible buttons in this container
                const buttons = container.querySelectorAll('button');
                const visibleButtons = Array.from(buttons).filter(b => b.offsetParent !== null);
                
                if (visibleButtons.length > 0) {
                    return {
                        containerLevel: i,
                        containerTag: container.tagName,
                        containerClass: container.className,
                        visibleButtonCount: visibleButtons.length,
                        buttons: visibleButtons.map(b => ({
                            class: b.className,
                            text: b.textContent.substring(0, 50),
                            ariaLabel: b.getAttribute('aria-label'),
                            dataTest: b.getAttribute('data-test'),
                            offsetHeight: b.offsetHeight,
                            offsetWidth: b.offsetWidth
                        }))
                    };
                }
                
                container = container.parentElement;
            }
            
            return { error: 'no visible button found near select' };
        """)
        print(f"[INFO] Result: {json.dumps(result, indent=2, ensure_ascii=False)}")

        # === STRATEGY 2: Search the entire document for elements that look like page size selector ===
        print("\n=== STRATEGY 2: Global Search for Page Size Selector ===")
        result2 = driver.execute_script("""
            // Look for elements with specific data attributes or ARIA labels
            const candidates = [];
            
            // 1. Elements with data-test or data-selector containing 'page', 'rows', 'size', 'limit'
            document.querySelectorAll('[data-test*="page"], [data-test*="rows"], [data-test*="size"], [data-test*="limit"], [data-test*="per"]').forEach(el => {
                if (el.offsetParent !== null) {
                    candidates.push({
                        type: 'data-test attribute',
                        dataTest: el.getAttribute('data-test'),
                        tag: el.tagName,
                        class: el.className.substring(0, 80),
                        text: el.textContent.substring(0, 30)
                    });
                }
            });
            
            // 2. Elements with aria-label containing relevant keywords
            document.querySelectorAll('[aria-label*="페이지"], [aria-label*="행"], [aria-label*="크기"], [aria-label*="보기"], [aria-label*="per"], [aria-label*="rows"]').forEach(el => {
                if (el.offsetParent !== null && candidates.length < 10) {
                    candidates.push({
                        type: 'aria-label attribute',
                        ariaLabel: el.getAttribute('aria-label'),
                        tag: el.tagName,
                        class: el.className.substring(0, 80),
                        text: el.textContent.substring(0, 30)
                    });
                }
            });
            
            // 3. Look for any visible dropdown or select-like structure
            document.querySelectorAll('[role="combobox"], [role="listbox"], [aria-haspopup]').forEach(el => {
                if (el.offsetParent !== null && candidates.length < 10) {
                    candidates.push({
                        type: 'role attribute',
                        role: el.getAttribute('role'),
                        ariaHasPopup: el.getAttribute('aria-haspopup'),
                        tag: el.tagName,
                        class: el.className.substring(0, 80),
                        text: el.textContent.substring(0, 50)
                    });
                }
            });
            
            return {
                candidateCount: candidates.length,
                candidates: candidates.slice(0, 15)
            };
        """)
        print(f"[INFO] Result: {json.dumps(result2, indent=2, ensure_ascii=False)}")

        # === STRATEGY 3: Try to trigger dropdown by direct JavaScript manipulation ===
        print("\n=== STRATEGY 3: Direct JavaScript Dropdown Trigger ===")
        trigger_result = driver.execute_script("""
            const select = document.querySelector('select.fds-field-select');
            if (!select) return { error: 'select not found' };
            
            // Try to find and click the associated button
            const parent = select.closest('.fds-field, [class*="select"]');
            if (!parent) return { error: 'parent not found' };
            
            // Look for any element with specific attributes that might trigger the dropdown
            const potentialTriggers = parent.querySelectorAll('[class*="button"], [role="button"], button, div[class*="trigger"], div[class*="dropdown"]');
            
            const results = [];
            for (let trigger of potentialTriggers) {
                if (trigger.offsetParent !== null) {
                    results.push({
                        tag: trigger.tagName,
                        class: trigger.className.substring(0, 100),
                        text: trigger.textContent.substring(0, 30),
                        role: trigger.getAttribute('role'),
                        offsetHeight: trigger.offsetHeight,
                        offsetWidth: trigger.offsetWidth,
                        canClick: true
                    });
                }
            }
            
            return {
                potentialTriggers: results
            };
        """)
        print(f"[INFO] Result: {json.dumps(trigger_result, indent=2, ensure_ascii=False)}")

        # === STRATEGY 4: Look at page size display text in the table header ===
        print("\n=== STRATEGY 4: Check Table for Page Size Indication ===")
        table_result = driver.execute_script("""
            const table = document.querySelector('[role="table"], .fds-table, table');
            if (!table) return { error: 'table not found' };
            
            // Look for pagination info like "1-10 of 58" or "showing 10 rows"
            const allText = table.parentElement.textContent;
            const pagination = allText.match(/\\d+-\\d+\\s+of\\s+\\d+|showing\\s+\\d+\\s+rows|per\\s+page:\\s+\\d+/gi);
            
            // Count actual visible rows
            const rows = table.querySelectorAll('tbody tr, [role="row"]');
            
            return {
                paginationText: pagination,
                visibleRowCount: rows.length,
                tableParent: table.parentElement.className.substring(0, 100)
            };
        """)
        print(f"[INFO] Table Info: {json.dumps(table_result, indent=2, ensure_ascii=False)}")

        print("\n[INFO] Analysis complete. Press Enter to close browser...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
