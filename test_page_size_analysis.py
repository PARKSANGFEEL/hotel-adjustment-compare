#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test: 페이지 사이즈 설정 문제 해결
목표: FDS dropdown에서 페이지 사이즈를 100으로 설정해서 테이블 행이 나타나는지 확인
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

        # Wait for payload
        for _ in range(20):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                break
            time.sleep(0.5)
        
        print("[INFO] Page loaded")
        time.sleep(1)

        # === 현재 테이블 상태 확인 ===
        print("\n=== 초기 상태 ===")
        initial_rows = driver.execute_script("""
            const rows = document.querySelectorAll('.fds-table tbody tr, [role="table"] [role="row"]');
            return rows.length;
        """)
        print(f"[INFO] 초기 테이블 행 개수: {initial_rows}")

        # === Native select 확인 ===
        print("\n=== Native Select 분석 ===")
        select_info = driver.execute_script("""
            const select = document.querySelector('select.fds-field-select');
            if (!select) return { error: 'select not found' };
            
            return {
                found: true,
                value: select.value,
                options: Array.from(select.options).map(o => o.value),
                visible: select.offsetParent !== null,
                display: window.getComputedStyle(select).display,
                parent_class: select.parentElement.className,
                parent_visible: select.parentElement.offsetParent !== null
            };
        """)
        print(json.dumps(select_info, indent=2, ensure_ascii=False))

        # === 전체 페이지에서 "100" 텍스트를 포함한 요소 찾기 ===
        print("\n=== 페이지에서 '100' 텍스트 찾기 ===")
        elements_with_100 = driver.execute_script("""
            const results = [];
            const allElements = document.querySelectorAll('*');
            
            for (let el of allElements) {
                const text = (el.textContent || '').trim();
                
                // 정확히 "100" 또는 "100개" 등을 포함
                if ((text === '100' || text.includes('100')) && el.offsetParent !== null) {
                    // 너무 큰 요소는 제외
                    if (el.children.length < 10) {
                        results.push({
                            tag: el.tagName,
                            class: el.className.substring(0, 100),
                            text: text.substring(0, 50),
                            offsetHeight: el.offsetHeight,
                            offsetWidth: el.offsetWidth,
                            visible: true
                        });
                    }
                }
            }
            
            return results.slice(0, 20);
        """)
        
        print(f"[INFO] '100'을 포함한 visible 요소 {len(elements_with_100)}개:")
        for el in elements_with_100[:10]:
            print(f"  - {el['tag']}.{el['class'][:50]}: {el['text']}")

        # === 페이지에서 모든 combobox/dropdown 찾기 ===
        print("\n=== 모든 Dropdown/Combobox 찾기 ===")
        dropdowns = driver.execute_script("""
            const results = [];
            
            // 1. role="combobox" 찾기
            document.querySelectorAll('[role="combobox"]').forEach(el => {
                if (el.offsetParent !== null) {
                    results.push({
                        type: 'combobox',
                        class: el.className.substring(0, 80),
                        text: el.textContent.substring(0, 50),
                        ariaLabel: el.getAttribute('aria-label'),
                        ariaExpanded: el.getAttribute('aria-expanded')
                    });
                }
            });
            
            // 2. 페이지 사이즈 관련 텍스트를 가진 버튼 찾기
            document.querySelectorAll('button, div[role="button"]').forEach(el => {
                const text = (el.textContent || '').toLowerCase();
                if ((text.includes('page') || text.includes('row') || text.includes('per')) && el.offsetParent !== null) {
                    results.push({
                        type: 'potential_page_size_button',
                        class: el.className.substring(0, 80),
                        text: el.textContent.substring(0, 50)
                    });
                }
            });
            
            return results.slice(0, 15);
        """)
        
        print(f"[INFO] 드롭다운/버튼 {len(dropdowns)}개 발견:")
        for dd in dropdowns[:10]:
            print(f"  - {dd['type']}: {dd.get('text', '')}")

        # === 현재 페이지 상태 스크린샷 ===
        print("\n=== 페이지 상태 ===")
        page_state = driver.execute_script("""
            return {
                current_url: window.location.href,
                select_visible: document.querySelector('select.fds-field-select') ? document.querySelector('select.fds-field-select').offsetParent !== null : false,
                table_exists: document.querySelector('.fds-table') ? true : false,
                table_rows: document.querySelectorAll('.fds-table tbody tr').length,
                body_height: document.body.scrollHeight
            };
        """)
        print(json.dumps(page_state, indent=2, ensure_ascii=False))

        print("\n[INFO] 분석 완료. 브라우저 창을 확인하고 Enter를 눌러주세요...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
