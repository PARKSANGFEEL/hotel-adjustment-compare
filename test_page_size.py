# -*- coding: utf-8 -*-
"""
Test: set page size to 100 on statements list.
Uses saved cookies in expedia_cookies.json.
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

        # 명세서 탭 클릭 (data-content-id="statements" 또는 텍스트 '명세서')
        try:
            tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-content-id="statements"]'))
            )
            tab.click()
            time.sleep(1)
            print("[OK] 명세서 탭 클릭")
        except Exception:
            try:
                tab = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='명세서']"))
                )
                tab.click()
                time.sleep(1)
                print("[OK] 명세서 탭 클릭 (텍스트)")
            except Exception as e:
                print(f"[WARN] 명세서 탭 클릭 실패: {e}")

        # wait payload
        for _ in range(20):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                break
            time.sleep(0.5)
        print("[INFO] payload ready")

        # find select
        # select 요소들 찾기
        selects = driver.find_elements(By.TAG_NAME, 'select')
        print(f"[INFO] selects on page: {len(selects)}")
        for i, s in enumerate(selects[:5]):
            opts = [o.get_attribute('value') for o in s.find_elements(By.TAG_NAME, 'option')]
            print(f"  [{i}] opts: {opts}")
        
        if not selects:
            print("[ERROR] no select found")
            return
        
        # 페이지네이션 select는 보통 마지막이거나 첫번째
        select_elem = None
        for select in selects:
            opts = [o.get_attribute('value') for o in select.find_elements(By.TAG_NAME, 'option')]
            if '100' in opts:
                select_elem = select
                print(f"[OK] pagination select found")
                break
        
        # === Deep Inspection ===
        print("\n=== SELECT ELEMENT ANALYSIS ===")
        result = driver.execute_script("""
            const select = document.querySelector('select.fds-field-select');
            if (!select) return { error: 'select not found' };
            
            return {
                tagName: select.tagName,
                className: select.className,
                isVisible: select.offsetParent !== null,
                offsetHeight: select.offsetHeight,
                offsetWidth: select.offsetWidth,
                parentTag: select.parentElement.tagName,
                parentClass: select.parentElement.className,
                grandparentTag: select.parentElement.parentElement.tagName,
                grandparentClass: select.parentElement.parentElement.className
            };
        """)
        print(f"[INFO] Select inspection: {json.dumps(result, ensure_ascii=False, indent=2)}")
        
        # === Search for visible dropdown UI ===
        print("\n=== SEARCHING FOR VISIBLE DROPDOWN UI ===")
        search_result = driver.execute_script("""
            // 1. Look for role="combobox" or aria-haspopup="listbox"
            const comboboxes = document.querySelectorAll('[role="combobox"], [aria-haspopup="listbox"]');
            console.log(`Found ${comboboxes.length} combobox/listbox elements`);
            
            // 2. Look for visible buttons near the select
            const select = document.querySelector('select.fds-field-select');
            if (!select) return { error: 'select not found' };
            
            const parent = select.closest('[class*="fds-field"]');
            if (!parent) return { error: 'fds-field parent not found' };
            
            const grandparent = parent.parentElement;
            if (!grandparent) return { error: 'grandparent not found' };
            
            // Get all siblings and descendants
            const allElements = grandparent.querySelectorAll('*');
            const visibleElements = Array.from(allElements).filter(el => el.offsetParent !== null);
            
            // Look for buttons, divs with role, or input-like elements
            const interactive = visibleElements.filter(el => {
                const tag = el.tagName.toLowerCase();
                const role = el.getAttribute('role');
                return tag === 'button' || tag === 'a' || 
                       role === 'button' || role === 'combobox' || role === 'listbox' ||
                       (tag === 'div' && el.getAttribute('aria-haspopup'));
            });
            
            return {
                parentVisible: parent.offsetParent !== null,
                grandparentVisible: grandparent.offsetParent !== null,
                totalVisibleInGrandparent: visibleElements.length,
                interactiveElements: interactive.map(el => ({
                    tag: el.tagName,
                    class: el.className.substring(0, 100),
                    role: el.getAttribute('role'),
                    text: el.textContent.substring(0, 50),
                    offsetHeight: el.offsetHeight,
                    offsetWidth: el.offsetWidth
                })).slice(0, 10)
            };
        """)
        print(f"[INFO] Dropdown search: {json.dumps(search_result, ensure_ascii=False, indent=2)}")
        
        if not select_elem:
            print("[ERROR] no select with option 100 found")
            return

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_elem)
        time.sleep(0.3)
        print("[INFO] scrolled to select")
        
        # 클릭 시도 1: JS click
        try:
            driver.execute_script("arguments[0].click();", select_elem)
            print("[OK] JS click successful")
            time.sleep(0.2)
        except Exception as e:
            print(f"[WARN] JS click failed: {e}")
        
        # 클릭 시도 2: 직접 option 100 선택
        try:
            opt100 = select_elem.find_element(By.XPATH, ".//option[@value='100']")
            driver.execute_script("arguments[0].selected = true;", opt100)
            print("[INFO] set option selected=true")
        except Exception as e:
            print(f"[WARN] option select failed: {e}")
        
        # 이벤트 발생 (더 강력한 방식)
        driver.execute_script("""
            const sel = arguments[0];
            // 1. 모든 option의 selected 초기화
            Array.from(sel.options).forEach(o => o.selected = false);
            // 2. 100 option 선택
            sel.value = '100';
            const opt100 = Array.from(sel.options).find(o => o.value === '100');
            if (opt100) opt100.selected = true;
            // 3. 다양한 이벤트 발생
            sel.dispatchEvent(new Event('mousedown', {bubbles:true}));
            sel.dispatchEvent(new Event('mouseup', {bubbles:true}));
            sel.dispatchEvent(new Event('click', {bubbles:true}));
            sel.dispatchEvent(new Event('input', {bubbles:true, composed:true}));
            sel.dispatchEvent(new Event('change', {bubbles:true, composed:true}));
            sel.dispatchEvent(new Event('blur', {bubbles:true}));
            // 4. 사용자 정의 이벤트도 발생
            if (window.jQuery) {
                jQuery(sel).trigger('change');
            }
            return 'events dispatched';
        """, select_elem)
        print("[INFO] dispatched comprehensive events")
        
        time.sleep(2)  # 더 오래 대기하여 페이지 업데이트 확인
        val = select_elem.get_attribute('value')
        print(f"[INFO] select value after: {val}")

        try:
            for _ in range(8):
                rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
                print(f"[DEBUG] row count: {len(rows)}")
                if len(rows) >= 60:
                    break
                time.sleep(1)
            print(f"[INFO] final row count: {len(rows)}")
        except Exception as e:
            print(f"[ERROR] row count check failed: {e}")

    finally:
        input("\n[INFO] 창을 닫으려면 Enter: ")
        driver.quit()

if __name__ == '__main__':
    main()
