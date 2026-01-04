# -*- coding: utf-8 -*-
"""
Expedia 로그인 및 명세서 다운로드 테스트
"""

import os
import sys
import time

# 환경변수 확인
username = os.environ.get('EXPEDIA_USERNAME')
password = os.environ.get('EXPEDIA_PASSWORD')

if not username or not password:
    print("환경변수가 설정되지 않았습니다.")
    print("다음 명령을 실행하세요:")
    print("  $env:EXPEDIA_USERNAME='gridinn'")
    print("  $env:EXPEDIA_PASSWORD='rmflemdls!2015'")
    sys.exit(1)

print(f"환경변수 확인:")
print(f"  EXPEDIA_USERNAME: {username}")
print(f"  EXPEDIA_PASSWORD: {'*' * len(password)}")

from expedia_downloader import ExpediaDownloader

print("\n[테스트] 로그인 및 명세서 다운로드")
print("="*80)

downloader = ExpediaDownloader()

try:
    downloader.setup_driver()
    print("\n드라이버 설정 완료. 로그인 시도...")
    
    success = downloader.login()
    
    if success:
        print("\n[SUCCESS] 로그인 성공!")
        
        # 명세서 페이지로 이동
        print("\n[명세서] 페이지로 이동 중...")
        downloader.navigate_to_statements()
        
        # 페이지 HTML 저장 (디버깅용)
        page_source = downloader.driver.page_source
        with open('page_source_statements.html', 'w', encoding='utf-8') as f:
            f.write(page_source)
        print("\n페이지 HTML 저장: page_source_statements.html")
        print(f"페이지 크기: {len(page_source)} bytes")
        
        # 테이블 구조 추출 (디버깅)
        import re
        table_match = re.search(r'<table[^>]*class="[^"]*fds-data-table[^"]*"[^>]*>.*?</table>', page_source, re.DOTALL)
        if table_match:
            table_html = table_match.group(0)[:5000]  # 처음 5000자
            with open('table_structure.html', 'w', encoding='utf-8') as f:
                f.write(table_html)
            print("테이블 구조 저장: table_structure.html")
        
        # 명세서 목록 조회
        print("\n[조회] 명세서 목록 조회 중...")
        statements = downloader.get_statement_list()
        
        if statements:
            print(f"\n[SUCCESS] {len(statements)}개 명세서 발견!")
            print("\n첫 5개 명세서:")
            for i, stmt in enumerate(statements[:5], 1):
                date_requested = stmt.get('dateRequested', 'N/A')
                payment_id = stmt.get('paymentRequestId', 'N/A')
                date_paid = stmt.get('datePaid', '미결제')
                print(f"  {i}. {date_requested} | {payment_id} | {date_paid}")
            
            # 최신 3개 다운로드
            print(f"\n[다운로드] 최신 3개 명세서 다운로드 시작...")
            downloaded = downloader.download_statements(limit=3)
            print(f"[SUCCESS] {downloaded}개 파일 다운로드 완료")
        else:
            print("\n[FAILED] 명세서를 찾을 수 없습니다")
        
        print("\n10초 후 브라우저를 닫습니다...")
        time.sleep(10)
        
    else:
        print("\n[FAILED] 로그인 실패")
        print("\n스크린샷 파일을 확인하세요:")
        print("  - login_step1.png")
        print("  - login_step2.png")
        print("  - after_login.png")
        print("  - error_screenshot.png")
        
        print("\n10초 후 브라우저를 닫습니다...")
        time.sleep(10)
    
    downloader.driver.quit()
    print("\n브라우저 종료")
    
except Exception as e:
    print(f"\n오류 발생: {e}")
    import traceback
    traceback.print_exc()
    
    if downloader.driver:
        downloader.driver.quit()

print("\n테스트 완료")
