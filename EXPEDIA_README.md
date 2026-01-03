# Expedia 명세서 자동 다운로드

Expedia Partner Central에서 명세서를 자동으로 다운로드하는 모듈입니다.

## 설치

### 1. 필수 패키지 설치

```bash
pip install selenium pandas openpyxl
```

### 2. Chrome 드라이버 설치

- Selenium은 Chrome 브라우저를 사용합니다
- Chrome 브라우저가 설치되어 있어야 합니다
- ChromeDriver는 자동으로 관리됩니다 (selenium 4.x 이상)

### 3. 환경변수 설정

#### Windows PowerShell (현재 세션만):
```powershell
$env:EXPEDIA_USERNAME='gridinn'
$env:EXPEDIA_PASSWORD='rmflemdls!2015'
```

#### Windows 시스템 환경변수 (영구 설정):
1. 시스템 속성 > 환경 변수
2. 새로 만들기:
   - 변수 이름: `EXPEDIA_USERNAME`
   - 변수 값: `gridinn`
3. 새로 만들기:
   - 변수 이름: `EXPEDIA_PASSWORD`
   - 변수 값: `rmflemdls!2015`

## 사용법

### 1. 테스트 실행

```bash
python test_expedia_downloader.py
```

테스트 메뉴:
- 1: 로그인만 테스트
- 2: 명세서 목록 조회 테스트
- 3: 다운로드 이력 테스트
- 4: 전체 테스트 실행

### 2. 독립 실행 (새로운 명세서만 다운로드)

```bash
python expedia_downloader.py
```

### 3. 날짜 범위 지정 다운로드

```bash
python expedia_downloader.py --start-date 2025-12-01 --end-date 2025-12-31
```

### 4. compare_sales.py와 함께 실행

#### 새로운 명세서만 다운로드 후 비교:
```bash
python compare_sales.py --download-expedia
```

#### 특정 기간 명세서 다운로드 후 비교:
```bash
python compare_sales.py --download-expedia --expedia-start-date 2025-12-01 --expedia-end-date 2025-12-31
```

#### 비교만 실행 (다운로드 안함):
```bash
python compare_sales.py
```

## 동작 방식

1. **로그인**: Expedia Partner Central 로그인
2. **명세서 조회**: 결제 > 송장 또는 명세서 찾기 페이지에서 목록 조회
3. **필터링**: 이미 다운로드한 명세서 제외 (또는 날짜 범위 필터)
4. **다운로드**: 각 지불 ID 클릭 → 목록 다운로드 버튼 클릭
5. **파일 이름 변경**: `익스피디아_(처리금액)_(결제날짜).xlsx`
6. **이력 저장**: `expedia_download_history.xlsx`에 다운로드 정보 기록

## 파일 구조

```
hotel-adjustment-compare/
├── expedia_downloader.py          # 메인 다운로더 모듈
├── test_expedia_downloader.py     # 테스트 스크립트
├── compare_sales.py                # 메인 비교 프로그램 (옵션 추가)
├── expedia_download_history.xlsx  # 다운로드 이력 (자동 생성)
├── ota-adjustment/                 # 다운로드 파일 저장 위치
│   ├── 익스피디아_123456_251229.xlsx
│   └── ...
└── temp_downloads/                 # 임시 다운로드 폴더 (자동 생성/삭제)
```

## 다운로드 이력 관리

- 파일: `expedia_download_history.xlsx`
- 컬럼:
  - `request_date`: 요청날짜
  - `tracking_number`: 추적번호
  - `payment_id`: 지불ID
  - `payment_date`: 결제날짜
  - `amount`: 처리금액
  - `filename`: 저장된 파일명
  - `download_time`: 다운로드 시각

## 트러블슈팅

### 로그인 실패
- 환경변수가 올바르게 설정되었는지 확인
- Expedia 계정 정보 확인
- 네트워크 연결 확인

### 브라우저가 열리지 않음
- Chrome 브라우저 설치 확인
- ChromeDriver 버전 확인 (자동 관리됨)

### 다운로드가 안됨
- 페이지 로딩 시간 확인 (느린 경우 time.sleep 값 조정)
- 명세서 페이지 구조 변경 확인
- 브라우저 디버깅 모드로 실행

### Permission Denied 오류
- 다운로드 폴더 권한 확인
- 이미 열려있는 엑셀 파일 닫기

## 주의사항

- **자동화 정책**: Expedia의 자동화 정책을 준수하세요
- **다운로드 간격**: 너무 빠른 요청은 차단될 수 있습니다
- **보안**: 환경변수에 비밀번호를 저장하므로 시스템 보안에 주의하세요
- **백업**: 다운로드 이력 파일을 정기적으로 백업하세요

## 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.
