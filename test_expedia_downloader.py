# -*- coding: utf-8 -*-
"""
Expedia 다운로더 테스트 스크립트

사용법:
    python test_expedia_downloader.py

테스트 항목:
1. 환경변수 확인
2. 로그인 테스트
3. 명세서 목록 조회 테스트
4. 다운로드 테스트 (실제 다운로드는 안함)
"""

import os
import sys
from pathlib import Path


def test_env_variables():
    """환경변수 확인"""
    print("\n" + "="*80)
    print("1. 환경변수 확인")
    print("="*80)
    
    username = os.environ.get('EXPEDIA_USERNAME')
    password = os.environ.get('EXPEDIA_PASSWORD')
    
    if username:
        print(f"✓ EXPEDIA_USERNAME: {username}")
    else:
        print("✗ EXPEDIA_USERNAME: 설정되지 않음")
    
    if password:
        print(f"✓ EXPEDIA_PASSWORD: {'*' * len(password)}")
    else:
        print("✗ EXPEDIA_PASSWORD: 설정되지 않음")
    
    if username and password:
        print("\n[SUCCESS] 환경변수 설정 완료")
        return True
    else:
        print("\n[FAILED] 환경변수를 설정하세요:")
        print("  Windows PowerShell:")
        print("    $env:EXPEDIA_USERNAME='gridinn'")
        print("    $env:EXPEDIA_PASSWORD='rmflemdls!2015'")
        print("\n  또는 시스템 환경변수에 추가하세요.")
        return False


def test_login_only():
    """로그인만 테스트"""
    print("\n" + "="*80)
    print("2. 로그인 테스트")
    print("="*80)
    
    try:
        from expedia_downloader import ExpediaDownloader
        
        downloader = ExpediaDownloader()
        downloader.setup_driver()
        
        success = downloader.login()
        
        if success:
            print("\n[SUCCESS] 로그인 성공!")
            print("5초 후 브라우저를 닫습니다...")
            import time
            time.sleep(5)
        else:
            print("\n[FAILED] 로그인 실패")
        
        downloader.driver.quit()
        return success
        
    except Exception as e:
        print(f"\n[ERROR] 테스트 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_statement_list():
    """명세서 목록 조회 테스트"""
    print("\n" + "="*80)
    print("3. 명세서 목록 조회 테스트")
    print("="*80)
    
    try:
        from expedia_downloader import ExpediaDownloader
        
        downloader = ExpediaDownloader()
        downloader.setup_driver()
        
        if not downloader.login():
            print("\n[FAILED] 로그인 실패")
            downloader.driver.quit()
            return False
        
        downloader.navigate_to_statements()
        statements = downloader.get_statement_list()
        
        if statements:
            print(f"\n[SUCCESS] {len(statements)}개 명세서 발견")
            print("\n처음 3개 명세서:")
            for i, stmt in enumerate(statements[:3], 1):
                print(f"\n{i}. 지불ID: {stmt['payment_id']}")
                print(f"   결제날짜: {stmt['payment_date']}")
                print(f"   처리금액: {stmt['amount']}")
        else:
            print("\n[FAILED] 명세서를 찾을 수 없습니다.")
        
        print("\n10초 후 브라우저를 닫습니다...")
        import time
        time.sleep(10)
        
        downloader.driver.quit()
        return len(statements) > 0
        
    except Exception as e:
        print(f"\n[ERROR] 테스트 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_download_history():
    """다운로드 이력 파일 테스트"""
    print("\n" + "="*80)
    print("4. 다운로드 이력 테스트")
    print("="*80)
    
    try:
        from expedia_downloader import ExpediaDownloader
        
        downloader = ExpediaDownloader()
        history_df = downloader.load_download_history()
        
        print(f"\n이력 파일: {downloader.history_file}")
        print(f"기록된 다운로드: {len(history_df)}개")
        
        if not history_df.empty:
            print("\n최근 3개 다운로드:")
            for i, row in history_df.tail(3).iterrows():
                print(f"\n{i+1}. 파일명: {row['filename']}")
                print(f"   지불ID: {row['payment_id']}")
                print(f"   결제날짜: {row['payment_date']}")
                print(f"   다운로드: {row['download_time']}")
        
        print("\n[SUCCESS] 이력 파일 정상")
        return True
        
    except Exception as e:
        print(f"\n[ERROR] 테스트 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """전체 테스트 실행"""
    print("="*80)
    print("Expedia 다운로더 테스트")
    print("="*80)
    
    # 필수 패키지 확인
    try:
        import selenium
        import pandas
        print("\n✓ 필수 패키지 설치 확인 완료")
    except ImportError as e:
        print(f"\n✗ 필수 패키지 누락: {e}")
        print("\n다음 명령으로 설치하세요:")
        print("  pip install selenium pandas openpyxl")
        return
    
    # 1. 환경변수 확인
    if not test_env_variables():
        return
    
    # 사용자 선택
    print("\n" + "="*80)
    print("테스트 선택")
    print("="*80)
    print("1. 로그인만 테스트")
    print("2. 명세서 목록 조회 테스트")
    print("3. 다운로드 이력 테스트")
    print("4. 전체 실행 (실제 다운로드 안함)")
    print("0. 종료")
    
    choice = input("\n선택 (0-4): ").strip()
    
    if choice == '1':
        test_login_only()
    elif choice == '2':
        test_statement_list()
    elif choice == '3':
        test_download_history()
    elif choice == '4':
        test_login_only()
        test_statement_list()
        test_download_history()
    elif choice == '0':
        print("\n테스트를 종료합니다.")
    else:
        print("\n잘못된 선택입니다.")


if __name__ == '__main__':
    main()
