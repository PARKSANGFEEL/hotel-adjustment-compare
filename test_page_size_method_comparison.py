# -*- coding: utf-8 -*-
"""
Test three methods to change select from 10 to 100.
Method 1: Selenium click on option
Method 2: sendKeys approach
Method 3: Direct API/form submission
"""
import json
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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

def get_row_count(driver):
    rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
    return len(rows)

def wait_for_row_change(driver, initial_count, max_wait=8):
    """Wait for row count to change from initial_count"""
    for i in range(max_wait * 2):
        current = get_row_count(driver)
        print(f"  [check {i}] row count: {current}")
        if current > initial_count + 10:
            print(f"[SUCCESS] rows increased from {initial_count} to {current}")
            return True
        time.sleep(0.5)
    print(f"[FAIL] rows unchanged: {current}")
    return False

def method1_selenium_click(driver, select_elem):
    """Method 1: Click the option element using Selenium"""
    print("\n=== METHOD 1: Selenium click on option element ===")
    initial_count = get_row_count(driver)
    print(f"Initial rows: {initial_count}")
    
    try:
        # 옵션 100 찾기
        option_100 = None
        for opt in select_elem.find_elements(By.TAG_NAME, 'option'):
            if opt.get_attribute('value') == '100':
                option_100 = opt
                break
        
        if not option_100:
            print("[ERROR] option 100 not found")
            return False
        
        # 옵션을 선택 가능하게 만들기 (일부 페이지에서는 옵션이 숨겨짐)
        driver.execute_script("arguments[0].style.display='block';", option_100)
        driver.execute_script("arguments[0].scrollIntoView();", option_100)
        
        # Selenium 클릭
        option_100.click()
        print("[OK] clicked option 100 with Selenium")
        
        # 결과 대기
        return wait_for_row_change(driver, initial_count)
        
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

def method2_selenium_select(driver, select_elem):
    """Method 2: Use Selenium's Select class"""
    print("\n=== METHOD 2: Selenium Select class ===")
    initial_count = get_row_count(driver)
    print(f"Initial rows: {initial_count}")
    
    try:
        select = Select(select_elem)
        select.select_by_value('100')
        print("[OK] selected '100' using Select class")
        
        return wait_for_row_change(driver, initial_count)
        
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

def method3_sendkeys(driver, select_elem):
    """Method 3: Click select and use sendKeys"""
    print("\n=== METHOD 3: Click select + sendKeys ===")
    initial_count = get_row_count(driver)
    print(f"Initial rows: {initial_count}")
    
    try:
        # Select 클릭
        select_elem.click()
        time.sleep(0.5)
        
        # 여러 번 down arrow 누르기 또는 직접 입력
        # 보통 10 → 25 → 50 → 100 순서이므로 3번 누르기
        select_elem.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
        select_elem.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
        select_elem.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
        select_elem.send_keys(Keys.RETURN)
        print("[OK] sent arrow keys and return")
        
        return wait_for_row_change(driver, initial_count)
        
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

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
            time.sleep(1)
            print("[OK] 명세서 탭 clicked")
        except Exception:
            print("[WARN] tab click failed")

        # wait payload
        for _ in range(20):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                break
            time.sleep(0.5)
        print("[INFO] payload ready")

        # find select with option 100
        selects = driver.find_elements(By.TAG_NAME, 'select')
        select_elem = None
        for select in selects:
            opts = [o.get_attribute('value') for o in select.find_elements(By.TAG_NAME, 'option')]
            if '100' in opts:
                select_elem = select
                break
        
        if not select_elem:
            print("[ERROR] no select with option 100 found")
            return
        
        print(f"[OK] found select, current value: {select_elem.get_attribute('value')}")
        
        # Try all three methods
        results = {}
        
        # Scroll to select
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_elem)
        time.sleep(0.5)
        
        # Method 1
        results['method1'] = method1_selenium_click(driver, select_elem)
        time.sleep(2)
        
        # Reload if needed to reset
        print("\n[INFO] Reloading page for next method...")
        driver.get(STATEMENTS_URL)
        time.sleep(3)
        try:
            tab = driver.find_element(By.CSS_SELECTOR, '[data-content-id="statements"]')
            tab.click()
            time.sleep(1)
        except:
            pass
        
        # Re-find select
        for _ in range(15):
            selects = driver.find_elements(By.TAG_NAME, 'select')
            select_elem = None
            for select in selects:
                opts = [o.get_attribute('value') for o in select.find_elements(By.TAG_NAME, 'option')]
                if '100' in opts:
                    select_elem = select
                    break
            if select_elem:
                break
            time.sleep(0.5)
        
        if select_elem:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_elem)
            time.sleep(0.5)
            
            # Method 2
            results['method2'] = method2_selenium_select(driver, select_elem)
            time.sleep(2)
        
        # Print summary
        print("\n" + "="*50)
        print("SUMMARY:")
        for method, result in results.items():
            status = "✓ SUCCESS" if result else "✗ FAILED"
            print(f"{method}: {status}")
        print("="*50)
        
        input("\nPress ENTER to close browser...")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
