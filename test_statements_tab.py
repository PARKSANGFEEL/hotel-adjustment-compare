#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test: 명세서 탭 클릭 후 페이지 제대로 로드되는지 확인
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
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body')))
        time.sleep(2)

        print("[INFO] 페이지 로드 시작")

        # === "명세서" 탭 찾기 및 클릭 ===
        print("\n[시도] '명세서' 탭 찾기...")
        
        # 방법 1: data-content-id="statements" 찾기
        try:
            tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-content-id="statements"]'))
            )
            print("[OK] '명세서' 탭 발견 (selector: data-content-id=statements)")
            tab.click()
            print("[OK] '명세서' 탭 클릭 완료")
            time.sleep(3)
        except Exception as e:
            print(f"[WARN] data-content-id 방법 실패: {e}")
            
            # 방법 2: aria-label 또는 텍스트로 찾기
            try:
                tabs = driver.find_elements(By.CSS_SELECTOR, '[role="tab"]')
                print(f"[INFO] 발견한 탭 {len(tabs)}개:")
                for i, t in enumerate(tabs):
                    text = t.text
                    aria = t.get_attribute('aria-label')
                    print(f"  [{i}] text='{text}' aria='{aria}'")
                    
                    if '명세서' in text or '명세서' in (aria or ''):
                        t.click()
                        print(f"[OK] 탭 {i}번 클릭 (명세서)")
                        time.sleep(3)
                        break
            except Exception as e2:
                print(f"[ERROR] 탭 클릭 실패: {e2}")

        # === 페이지 로드 대기 ===
        print("\n[대기] JavaScript 데이터 로드 중...")
        for i in range(30):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                print("[OK] JavaScript 페이로드 발견!")
                time.sleep(2)
                break
            print(f"  {i+1}/30...")
            time.sleep(1)

        # === 테이블 상태 확인 ===
        print("\n=== 테이블 상태 ===")
        table_info = driver.execute_script("""
            return {
                table_exists: document.querySelector('.fds-table') ? true : false,
                table_rows: document.querySelectorAll('.fds-table tbody tr').length,
                select_value: document.querySelector('select.fds-field-select')?.value || 'N/A',
                select_visible: document.querySelector('select.fds-field-select')?.offsetParent !== null
            };
        """)
        print(json.dumps(table_info, indent=2, ensure_ascii=False))

        if table_info['table_rows'] > 0:
            print(f"\n[SUCCESS] 테이블에 {table_info['table_rows']}개 행이 있습니다!")
        else:
            print("\n[PROBLEM] 테이블이 여전히 비어있습니다.")
            
            # 추가 디버깅
            print("\n[디버깅] 페이지의 모든 select 요소:")
            selects = driver.execute_script("""
                return Array.from(document.querySelectorAll('select')).map(s => ({
                    value: s.value,
                    options: Array.from(s.options).map(o => o.value),
                    visible: s.offsetParent !== null
                }));
            """)
            print(json.dumps(selects, indent=2, ensure_ascii=False))

        print("\n[INFO] 분석 완료. 브라우저를 확인하고 Enter를 눌러주세요...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
