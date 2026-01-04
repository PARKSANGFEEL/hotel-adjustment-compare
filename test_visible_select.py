#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test: Visible select 찾아서 100으로 변경 후 테이블 확인
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
        return False
    cookies = json.loads(COOKIES_FILE.read_text())
    driver.get("https://www.expediapartnercentral.com/Account/Logon?signedOff=true")
    time.sleep(1)
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
        time.sleep(2)

        # 명세서 탭 클릭
        try:
            tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-content-id="statements"]'))
            )
            tab.click()
            print("[OK] 명세서 탭 클릭")
            time.sleep(3)
        except:
            pass

        # 페이로드 대기
        for i in range(30):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                print("[OK] 페이로드 로드")
                time.sleep(2)
                break
            time.sleep(1)

        # === Visible select 찾기 ===
        print("\n=== Visible select 찾기 ===")
        selects = driver.find_elements(By.CSS_SELECTOR, 'select.fds-field-select')
        print(f"[INFO] {len(selects)}개 select 발견")
        
        visible_select = None
        for idx, sel in enumerate(selects):
            parent_visible = driver.execute_script("return arguments[0].parentElement.offsetParent !== null", sel)
            print(f"  [{idx}] parent_visible: {parent_visible}")
            if parent_visible:
                visible_select = sel
                print(f"[OK] Visible select 사용: index {idx}")
                break
        
        if not visible_select:
            print("[ERROR] Visible select를 찾지 못함!")
            return

        # 초기 테이블 상태
        initial_table = driver.execute_script("""
            return {
                rows: document.querySelectorAll('.fds-table tbody tr, [role="table"] [role="row"]').length
            };
        """)
        print(f"\n[BEFORE] 테이블 행 개수: {initial_table['rows']}")

        # === 값 변경 ===
        print("\n[시도] select 값을 100으로 변경...")
        driver.execute_script("""
            const select = arguments[0];
            select.value = '100';
            
            // 모든 가능한 이벤트 발생
            ['input', 'change', 'blur', 'click'].forEach(eventType => {
                const event = new Event(eventType, { bubbles: true, cancelable: true });
                select.dispatchEvent(event);
            });
            
            // 부모에도 change 이벤트
            const changeEvent = new Event('change', { bubbles: true });
            select.parentElement.dispatchEvent(changeEvent);
        """, visible_select)
        
        print("[OK] select 값 변경 완료")
        time.sleep(2)

        # 값 확인
        new_value = visible_select.get_attribute('value')
        print(f"[INFO] 변경 후 select 값: {new_value}")

        # 테이블 다시 확인
        after_table = driver.execute_script("""
            return {
                rows: document.querySelectorAll('.fds-table tbody tr, [role="table"] [role="row"]').length,
                table_exists: document.querySelector('.fds-table') !== null
            };
        """)
        print(f"\n[AFTER] 테이블 행 개수: {after_table['rows']}")
        print(f"[INFO] 테이블 존재: {after_table['table_exists']}")

        if after_table['rows'] > initial_table['rows']:
            print(f"\n[SUCCESS] 테이블 행이 {initial_table['rows']} → {after_table['rows']}로 증가!")
        elif after_table['rows'] == 0:
            print("\n[PROBLEM] 테이블이 여전히 비어있습니다.")
            print("[정보] 이는 페이지가 완전히 로드되지 않았거나 테이블이 렌더링되지 않은 것입니다.")
        else:
            print(f"\n[INFO] 테이블 행 개수가 동일합니다: {after_table['rows']}")

        print("\n[INFO] 브라우저를 확인하세요. Enter를 눌러 종료...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
