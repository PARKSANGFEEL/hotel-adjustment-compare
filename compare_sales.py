import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 파일 경로 설정
dir_base = os.path.dirname(os.path.abspath(__file__))
file_all_customers = os.path.join(dir_base, '전체고객 목록_20251230.xlsx')
directory_ota = os.path.join(dir_base, 'ota-adjustment')

# 전체고객 엑셀 파일 로드
df_all = pd.read_excel(file_all_customers)

# Remittances로 시작하는 파일 목록
ota_files = [f for f in os.listdir(directory_ota) if f.startswith('Remittances') and f.endswith('.xlsx')]

# 아고다 매출 데이터 통합
df_ota = pd.DataFrame()
for file in ota_files:
    path = os.path.join(directory_ota, file)
    temp_df = pd.read_excel(path)
    df_ota = pd.concat([df_ota, temp_df], ignore_index=True)



# 컬럼명 매핑 (수동 지정)
col_name_all = '고객명'  # D열
col_price_all_1 = '객실료'  # K열
col_price_all_2 = '합계'    # M열

# Remittances 엑셀의 이름, 금액 컬럼명 추정 (수정 필요시 아래 변수명 변경)
# 이름 컬럼: 4번째 컬럼(D열)
col_name_ota = df_ota.columns[3]
# 금액 컬럼: Remittances의 G열, H열 등 여러 컬럼을 모두 비교
ota_price_cols = [col for col in df_ota.columns if any(x in col for x in ['금액', 'Amount', '금액', '금액', '금액']) or col in ['G', 'H'] or col.startswith('Unnamed')]
if not ota_price_cols:
    # Remittances의 7,8번째 컬럼(G,H열)로 강제 지정
    ota_price_cols = [df_ota.columns[6], df_ota.columns[7]]

# 전체고객 엑셀 파일에 결과 표시를 위한 워크북 로드
wb = load_workbook(file_all_customers)
ws = wb.active

# 비교로그 시트 생성(기존 있으면 삭제)
if '비교로그' in wb.sheetnames:
    del wb['비교로그']
log_ws = wb.create_sheet('비교로그')
log_ws.append(['고객명', '전체매출 행번호', '전체매출 가격', 'Remittances 파일명', 'Remittances 행번호', '비교 Remittances 가격'])

# Remittances 파일별 행 오프셋 기록
ota_file_map = []  # (파일명, 데이터프레임, 시작행)
ota_row_offset = 0
df_ota = pd.DataFrame()
for file in ota_files:
    path = os.path.join(directory_ota, file)
    temp_df = pd.read_excel(path)
    ota_file_map.append((file, temp_df, ota_row_offset))
    ota_row_offset += len(temp_df)
    df_ota = pd.concat([df_ota, temp_df], ignore_index=True)

# 색상 스타일 정의
fill_yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
font_red = Font(color='FF0000')


# 이름과 가격 비교 및 색상 처리
# Remittances 이름 매칭 카운트 추적용
matched_remit_names = {}

for idx, row in df_all.iterrows():
    # 거래처가 아고다인 경우에만 비교
    vendor = str(row.get('거래처', '')).strip()
    if vendor != '아고다':
        continue
    name = str(row.get(col_name_all, '')).strip()
    price1 = str(row.get(col_price_all_1, '')).replace(',', '').strip()
    price2 = str(row.get(col_price_all_2, '')).replace(',', '').strip()
    # 금액을 float으로 변환
    try:
        price1_f = float(price1)
    except:
        price1_f = None
    try:
        price2_f = float(price2)
    except:
        price2_f = None
    # 이름이 일치하는 Remittances 행 찾기
    match = df_ota[df_ota[col_name_ota].astype(str).str.strip() == name]
    if not match.empty:
        price_match = False
        log_info = None
        for abs_idx, match_row in match.iterrows():
            for price_col in ota_price_cols:
                try:
                    ota_price_f = float(str(match_row[price_col]).replace(',', '').strip())
                except:
                    continue
                if (price1_f is not None and ota_price_f == price1_f) or (price2_f is not None and ota_price_f == price2_f):
                    price_match = True
                    # Remittances 매칭 카운트 기록
                    matched_remit_names[name] = matched_remit_names.get(name, 0) + 1
                    break
                # 로그 정보 저장 (조건 불일치 시)
                if not price_match and log_info is None:
                    for fname, df, offset in ota_file_map:
                        if offset <= abs_idx < offset + len(df):
                            file_row = abs_idx - offset + 2  # 2: 엑셀 헤더 보정
                            log_info = [name, idx+2, price1, fname, file_row, str(match_row[price_col])]
                            break
            if price_match:
                break
        if price_match:
            for cell in ws[idx+2]:
                cell.fill = fill_yellow
        else:
            # Remittances에 이름이 있지만, 이미 매칭된 횟수 이상이면 표시 없음
            if matched_remit_names.get(name, 0) > 0:
                matched_remit_names[name] -= 1
                continue
            for cell in ws[idx+2]:
                cell.font = font_red
            if log_info:
                log_ws.append(log_info)
    # Remittances에 없는 고객명은 아무 표시도 하지 않음

# 결과 저장
result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
wb.save(result_path)
print(f'완료: {result_path}에 저장됨')
