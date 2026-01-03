# -*- coding: utf-8 -*-
"""
Expedia 쿠키를 사용한 명세서 다운로드 테스트
"""

import os
import sys
import time
from pathlib import Path

from expedia_downloader import ExpediaDownloader

print("\n[테스트] 쿠키를 사용한 명세서 다운로드")
print("="*80)

downloader = ExpediaDownloader()

try:
    downloader.setup_driver()
    print("\n드라이버 설정 완료.")
    
    # 로그인 (쿠키 우선, 실패 시 수동 로그인)
    if downloader.login():
        print("[OK] 로그인 완료")
        
        # 명세서 페이지로 이동
        print("\n[명세서] 페이지로 이동 중...")
        downloader.navigate_to_statements()
        time.sleep(3)
        
        # 명세서 목록 조회
        print("\n[명세서] 목록 조회 중...")
        statements = downloader.get_statement_list()
        
        if statements:
            print(f"\n[OK] 총 {len(statements)}개 명세서 발견")
            
            # 첫 5개 출력
            print("\n첫 5개 명세서:")
            for i, stmt in enumerate(statements[:5], 1):
                date_requested = stmt.get('dateRequested', '-')
                payment_id = stmt.get('paymentRequestId', '-')
                date_paid = stmt.get('datePaid', '-')
                print(f"  {i}. {date_requested} | {payment_id} | {date_paid}")
            
            # 기간 설정 (예: 최근 1년, 또는 특정 기간)
            # date_from = '2024-01-01'  # 시작일 지정 (선택)
            # date_to = '2026-01-02'     # 종료일 지정 (선택)
            # downloader.download_statements(date_from=date_from, date_to=date_to)
            
            # 기본값 사용 (최근 1년치)
            downloader.download_statements()
            
            # 다운로드 폴더 확인
            files = list(downloader.download_dir.glob('익스피디아_*.csv'))
            if files:
                print(f"\n다운로드된 파일 ({len(files)}개):")
                for f in sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)[:10]:
                    print(f"  - {f.name}")
        else:
            print("[ERROR] 명세서를 찾을 수 없음")
    else:
        print("[ERROR] 로그인 실패")

except Exception as e:
    print(f"\n[ERROR] 오류 발생: {e}")
    import traceback
    traceback.print_exc()

finally:
    if downloader.driver:
        print("\n드라이버 종료...")
        downloader.driver.quit()
    print("[완료]")
