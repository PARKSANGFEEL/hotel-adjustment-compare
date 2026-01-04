#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test: 실제 visible한 dropdown UI 찾기
Native select의 부모 주변에서 모든 visible 요소를 상세 분석
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

        # === Native select 찾기 ===
        select_elem = driver.find_element(By.CSS_SELECTOR, 'select.fds-field-select')
        print("[OK] Native select 발견")

        # === Select의 부모 트리 상세 분석 ===
        print("\n=== Select 부모 트리 분석 ===")
        parent_structure = driver.execute_script("""
            const select = document.querySelector('select.fds-field-select');
            const result = [];
            
            let current = select;
            for (let i = 0; i < 8; i++) {
                if (!current) break;
                
                const info = {
                    level: i,
                    tag: current.tagName,
                    class: current.className,
                    id: current.id || 'none',
                    visible: current.offsetParent !== null,
                    offsetHeight: current.offsetHeight,
                    offsetWidth: current.offsetWidth,
                    children: current.children.length
                };
                result.push(info);
                
                // 이 레벨의 모든 child를 나열
                const children = Array.from(current.children).map((child, idx) => ({
                    index: idx,
                    tag: child.tagName,
                    class: child.className.substring(0, 80),
                    visible: child.offsetParent !== null,
                    text: child.textContent.substring(0, 50)
                }));
                
                console.log(`Level ${i} (${current.tagName}.${current.className}):`);
                console.log(`  Children: ${children.length}`);
                children.forEach(c => {
                    console.log(`    [${c.index}] ${c.tag} (visible: ${c.visible}) - ${c.class}`);
                });
                
                current = current.parentElement;
            }
            
            return result;
        """)
        
        for level in parent_structure:
            print(f"  Level {level['level']}: {level['tag']}.{level['class'][:50]} (visible: {level['visible']})")

        # === Visible siblings 찾기 (같은 부모 아래의 다른 요소들) ===
        print("\n=== Select의 부모의 모든 자식들 ===")
        siblings = driver.execute_script("""
            const select = document.querySelector('select.fds-field-select');
            const parent = select.parentElement; // LABEL
            const grandparent = parent.parentElement;
            
            // Grandparent의 모든 자식 나열
            const result = [];
            Array.from(grandparent.children).forEach((child, idx) => {
                const info = {
                    index: idx,
                    tag: child.tagName,
                    class: child.className.substring(0, 100),
                    visible: child.offsetParent !== null,
                    text: child.textContent.substring(0, 80),
                    role: child.getAttribute('role'),
                    offsetHeight: child.offsetHeight,
                    offsetWidth: child.offsetWidth
                };
                result.push(info);
            });
            
            return {
                grandparent_tag: grandparent.tagName,
                grandparent_class: grandparent.className,
                children_count: grandparent.children.length,
                children: result
            };
        """)
        
        print(f"Grandparent: {siblings['grandparent_tag']}.{siblings['grandparent_class']}")
        print(f"자식 개수: {siblings['children_count']}")
        print("\n자식 요소들:")
        for child in siblings['children']:
            visible_mark = "✓" if child['visible'] else "✗"
            print(f"  [{visible_mark}] {child['tag']}.{child['class'][:60]}")
            if child['text']:
                print(f"      Text: {child['text'][:60]}")

        # === 페이지의 모든 select 주변 요소 검사 ===
        print("\n=== Select 주변 모든 요소 상세 검사 ===")
        all_near_elements = driver.execute_script("""
            const select = document.querySelector('select.fds-field-select');
            const parent = select.parentElement;
            const grandparent = parent.parentElement;
            const greatgrandparent = grandparent.parentElement;
            
            // 모든 부모의 자식들을 모두 수집
            const allElements = [];
            
            [parent, grandparent, greatgrandparent].forEach((p, pIdx) => {
                if (!p) return;
                Array.from(p.querySelectorAll('*')).forEach(el => {
                    if (el.offsetParent !== null && el !== select) { // visible이고 select 아님
                        allElements.push({
                            tag: el.tagName,
                            class: el.className.substring(0, 100),
                            text: el.textContent.substring(0, 50),
                            role: el.getAttribute('role'),
                            ariaLabel: el.getAttribute('aria-label'),
                            ariaExpanded: el.getAttribute('aria-expanded'),
                            onclick: !!el.onclick,
                            dataAttrs: Array.from(el.attributes)
                                .filter(a => a.name.startsWith('data-'))
                                .map(a => `${a.name}=${a.value}`)
                                .slice(0, 3)
                        });
                    }
                });
            });
            
            return allElements.slice(0, 30);
        """)
        
        print(f"발견된 visible 요소: {len(all_near_elements)}개\n")
        for i, el in enumerate(all_near_elements[:20]):
            print(f"  [{i}] {el['tag']}.{el['class'][:70]}")
            if el['text']:
                print(f"       Text: '{el['text']}'")
            if el['ariaLabel']:
                print(f"       aria-label: '{el['ariaLabel']}'")
            if el['dataAttrs']:
                print(f"       data: {el['dataAttrs']}")

        print("\n\n[중요] 브라우저를 열어서 개발자 도구로 다음을 확인해주세요:")
        print("1. Native <select> 요소 우클릭 → Inspect")
        print("2. 부모 요소들을 따라가며 어느 부모까지 visible인지 확인")
        print("3. 부모 또는 형제 요소 중에 버튼/div 형태의 dropdown UI 찾기")
        print("\nEnter를 눌러주세요...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
