#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test: 페이지 전체에서 페이지 사이즈 관련 UI 찾기
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
            time.sleep(3)
        except:
            pass

        # 페이로드 대기
        for i in range(30):
            if 'statementsAndInvoicesPayload' in driver.page_source:
                time.sleep(2)
                break
            time.sleep(1)

        # === 페이지 전체에서 "10", "25", "50", "100" 버튼/링크 찾기 ===
        print("=== 페이지 전체에서 페이지 사이즈 버튼 찾기 ===\n")
        
        size_buttons = driver.execute_script("""
            const result = [];
            
            // 모든 가능한 페이지 사이즈 값
            const sizes = ['10', '25', '50', '100'];
            
            // 모든 버튼, 링크, div[role="button"] 검사
            const candidates = document.querySelectorAll('button, a, [role="button"], span, div');
            
            for (let el of candidates) {
                const text = (el.textContent || '').trim();
                
                // 이 요소가 정확히 "10" 또는 "100" 등을 포함하면서 visible한지?
                for (let size of sizes) {
                    if (text === size || text === size + '개' || text === size + 'items') {
                        if (el.offsetParent !== null) { // visible
                            result.push({
                                tag: el.tagName,
                                class: el.className.substring(0, 100),
                                text: text,
                                role: el.getAttribute('role'),
                                ariaLabel: el.getAttribute('aria-label'),
                                position: {
                                    top: el.getBoundingClientRect().top,
                                    left: el.getBoundingClientRect().left
                                }
                            });
                        }
                    }
                }
            }
            
            return result;
        """)
        
        print(f"발견된 페이지 사이즈 버튼: {len(size_buttons)}개\n")
        for btn in size_buttons:
            print(f"  {btn['tag']}.{btn['class'][:60]}")
            print(f"    Text: {btn['text']}")
            print(f"    Position: ({btn['position']['top']:.0f}, {btn['position']['left']:.0f})")
            if btn['ariaLabel']:
                print(f"    aria-label: {btn['ariaLabel']}")
            print()

        # === 모든 select 요소 찾기 ===
        print("\n=== 페이지의 모든 SELECT 요소 ===\n")
        all_selects = driver.execute_script("""
            const result = [];
            
            document.querySelectorAll('select').forEach((sel, idx) => {
                const options = Array.from(sel.options).map(o => ({
                    value: o.value,
                    text: o.text
                }));
                
                result.push({
                    index: idx,
                    name: sel.name || 'unnamed',
                    value: sel.value,
                    visible: sel.offsetParent !== null,
                    options: options,
                    parent_class: sel.parentElement.className,
                    parent_visible: sel.parentElement.offsetParent !== null
                });
            });
            
            return result;
        """)
        
        for sel in all_selects:
            vis = "✓" if sel['visible'] else "✗"
            parent_vis = "✓" if sel['parent_visible'] else "✗"
            print(f"  [{vis}] SELECT (parent: {parent_vis}) - name={sel['name']}, value={sel['value']}")
            print(f"      Parent: {sel['parent_class'][:80]}")
            print(f"      Options: {[o['value'] for o in sel['options']]}")
            print()

        # === 페이지 하단의 pagination 요소 찾기 ===
        print("\n=== 페이지 하단의 Pagination 요소 ===\n")
        pagination = driver.execute_script("""
            // 테이블 아래를 보기
            const table = document.querySelector('[role="table"], .fds-table');
            if (!table) return { error: 'no table' };
            
            const tableParent = table.parentElement;
            const nextSiblings = [];
            
            let next = tableParent.nextElementSibling;
            for (let i = 0; i < 5 && next; i++) {
                nextSiblings.push({
                    tag: next.tagName,
                    class: next.className.substring(0, 100),
                    text: next.textContent.substring(0, 100),
                    visible: next.offsetParent !== null,
                    children: next.children.length
                });
                next = next.nextElementSibling;
            }
            
            // 또는 pagination nav 찾기
            const navPagination = document.querySelector('nav.fds-pagination, [class*="pagination"]');
            
            return {
                nextSiblings,
                navPagination: navPagination ? {
                    class: navPagination.className,
                    visible: navPagination.offsetParent !== null,
                    HTML: navPagination.innerHTML.substring(0, 200)
                } : null
            };
        """)
        
        print(json.dumps(pagination, indent=2, ensure_ascii=False))

        print("\n\n[INFO] 분석 완료. Enter를 눌러주세요...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
