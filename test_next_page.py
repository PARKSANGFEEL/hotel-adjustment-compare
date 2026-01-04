# -*- coding: utf-8 -*-
"""
Test: click next-page arrow on statements list.
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

        # 명세서 탭 클릭
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

        # 현재 페이지 정보
        current_info = driver.execute_script("return document.body.innerText.slice(0, 200);")
        print(f"[INFO] current page snippet: {current_info!r}")

        time.sleep(1)
        btn = driver.execute_script("""
            const icon = document.querySelector('button .fds-icon-name-arrow-forward-ios');
            if (icon) {
                const b = icon.closest('button');
                if (b && !b.disabled && b.getAttribute('aria-disabled') !== 'true') {
                    b.click();
                    return true;
                }
            }
            const candidates = Array.from(document.querySelectorAll('button, a'));
            for (const el of candidates) {
                const disabled = el.disabled || el.getAttribute('aria-disabled') === 'true';
                if (disabled) continue;
                const hasIcon = el.querySelector('.fds-icon-name-arrow-forward-ios, .fds-icon-name-chevron-right, .fds-icon-name-arrow-forward, .fds-icon-name-arrow-right');
                const hasUse = el.querySelector('use[href*="arrow-forward"], use[xlink\\:href*="arrow-forward"], use[href*="chevron-right"], use[xlink\\:href*="chevron-right"]');
                const aria = (el.getAttribute('aria-label') || '').toLowerCase();
                if (hasIcon || hasUse || aria.includes('next') || aria.includes('다음')) {
                    el.click();
                    return true;
                }
            }
            return false;
        """)
        print(f"[INFO] next button clicked: {btn}")
        time.sleep(1)
        # simple check: any page indicator text changes?
        page_info = driver.execute_script("""
            const el = document.querySelector('[aria-live]') || document.body;
            return el.innerText.slice(0,120);
        """)
        print(f"[INFO] page info snippet: {page_info!r}")

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
