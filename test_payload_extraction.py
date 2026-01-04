#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test: 페이로드에서 데이터 추출 (테이블 없이)
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

def extract_payload(page_source):
    """페이로드 JSON 추출"""
    start_marker = 'statementsAndInvoicesPayload: '
    start_idx = page_source.find(start_marker)
    
    if start_idx == -1:
        return None
    
    start_idx += len(start_marker)
    
    # JSON 객체 끝 찾기
    brace_count = 0
    in_string = False
    escape_next = False
    end_idx = start_idx
    
    for i in range(start_idx, len(page_source)):
        char = page_source[i]
        
        if escape_next:
            escape_next = False
            continue
        
        if char == '\\':
            escape_next = True
            continue
        
        if char == '"' and not escape_next:
            in_string = not in_string
            continue
        
        if not in_string:
            if char == '{':
                brace_count += 1
            elif char == '}':
                brace_count -= 1
                if brace_count == 0:
                    end_idx = i + 1
                    break
    
    json_str = page_source[start_idx:end_idx]
    try:
        return json.loads(json_str)
    except:
        return None

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
                print("[OK] 페이로드 발견!")
                time.sleep(2)
                break
            time.sleep(1)

        # === 페이로드에서 데이터 추출 ===
        print("\n[추출] 페이로드 JSON 데이터...")
        payload = extract_payload(driver.page_source)
        
        if not payload:
            print("[ERROR] 페이로드 추출 실패")
            return
        
        print(f"[OK] 페이로드 파싱 성공")
        print(f"[INFO] 페이로드 키: {list(payload.keys())}")
        
        # statements 추출
        if 'statements' in payload and 'paymentList' in payload['statements']:
            payments = payload['statements']['paymentList']
            print(f"\n[OK] {len(payments)}개 명세서 발견!")
            
            print("\n첫 3개 명세서:")
            for i, p in enumerate(payments[:3]):
                print(f"  [{i}]")
                print(f"    paymentRequestId: {p.get('paymentRequestId')}")
                print(f"    invoiceId: {p.get('invoiceId')}")
                print(f"    datePaid: {p.get('datePaid')}")
                print(f"    paymentRequestFilePath: {p.get('paymentRequestFilePath')[:100] if p.get('paymentRequestFilePath') else 'N/A'}")
        
        # === 이제 페이지에 테이블이 없어도 다운로드 가능! ===
        print("\n[결론]")
        print("페이로드에서 모든 데이터를 추출할 수 있습니다.")
        print("테이블을 클릭할 필요 없이 paymentRequestFilePath를 직접 사용하면 됩니다!")
        
        print("\n[INFO] Enter를 눌러주세요...")
        input()

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
