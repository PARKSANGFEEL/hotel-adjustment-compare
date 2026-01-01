
# ...기존 코드...

# 파일 맨 아래에 print_result_rows 정의 및 호출

# ...기존 코드 맨 아래에 추가...
def write_ratio_to_result_log(ratio, row_idx):
    from openpyxl import load_workbook
    result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
    wb = load_workbook(result_path)
    if '비교로그' not in wb.sheetnames:
        print('[비교로그] 시트가 없습니다.')
        return
    ws = wb['비교로그']
    # row_idx는 1부터 시작(헤더 포함), G열(7번째 열)에 기록
    ws.cell(row=row_idx, column=7).value = ratio
    wb.save(result_path)
    print(f"[진단-비교로그] {row_idx}행 G열에 비율 {ratio} 기록 완료")
import re

def normalize(val):
    if val is None:
        return ''
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', str(val)).replace('.0','').strip().lower()
def print_peter_ludwig_log():
    from openpyxl import load_workbook
    result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
    wb = load_workbook(result_path, data_only=True)
    if '비교로그' not in wb.sheetnames:
        print('[비교로그] 시트가 없습니다.')
        return
    ws = wb['비교로그']
    # print('[비교로그] Peter Ludwig 관련 행:')
    for row in ws.iter_rows(values_only=True):
        if row and any('Peter Ludwig' in str(cell) for cell in row):
            print(row)

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from collections import defaultdict
import re

# 파일 경로 설정
dir_base = os.path.dirname(os.path.abspath(__file__))

directory_ota = os.path.join(dir_base, 'ota-adjustment')

# 결과 엑셀 파일에서만 데이터 로드

# 결과파일이 없으면 전체고객 목록 파일을 복사해서 생성
result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
if not os.path.exists(result_path):
    import glob
    import shutil
    all_list = sorted(glob.glob(os.path.join(dir_base, '전체고객 목록_*.xlsx')))
    if not all_list:
        raise FileNotFoundError('전체고객 목록_*.xlsx 파일이 없습니다. 결과파일을 생성할 수 없습니다.')
    latest_all = all_list[-1]
    shutil.copy(latest_all, result_path)
df_all = pd.read_excel(result_path, sheet_name=0)

# 컬럼명 매핑 (자동 추출)
def find_col(cols, keyword):
    for c in cols:
        if keyword in c:
            return c
    return None

# OTA번호 컬럼이 있으면 항상 문자열로 변환(.0 제거)
col_ota_no = find_col(df_all.columns, 'OTA')
if col_ota_no:
    df_all[col_ota_no] = df_all[col_ota_no].astype(str).str.replace('.0', '', regex=False).str.strip()

# Remittances로 시작하는 파일 목록
ota_files = [f for f in os.listdir(directory_ota) if f.startswith('Remittances') and f.endswith('.xlsx')]

# 아고다 매출 데이터 통합
df_ota = pd.DataFrame()
for file in ota_files:
    path = os.path.join(directory_ota, file)
    temp_df = pd.read_excel(path)
    df_ota = pd.concat([df_ota, temp_df], ignore_index=True)




# 컬럼명 매핑 (자동 추출)
def find_col(cols, keyword):
    for c in cols:
        if keyword in c:
            return c
    return None

col_name_all = find_col(df_all.columns, '고객')
col_price_all_1 = find_col(df_all.columns, '객실')
col_price_all_2 = find_col(df_all.columns, '합계')
col_vendor = find_col(df_all.columns, '거래처')
col_ota_no = find_col(df_all.columns, 'OTA')

"""
# print(f"[DEBUG] 컬럼명 매핑: 고객명={col_name_all}, 객실료={col_price_all_1}, 합계={col_price_all_2}, 거래처={col_vendor}, OTA번호={col_ota_no}")
"""

# Remittances 엑셀의 이름, 금액 컬럼명 추정 (수정 필요시 아래 변수명 변경)
# 이름 컬럼: 4번째 컬럼(D열)
col_name_ota = df_ota.columns[3]
# 금액 컬럼: Remittances의 G열, H열 등 여러 컬럼을 모두 비교
ota_price_cols = [col for col in df_ota.columns if any(x in col for x in ['금액', 'Amount', '금액', '금액', '금액']) or col in ['G', 'H'] or col.startswith('Unnamed')]
if not ota_price_cols:
    # Remittances의 7,8번째 컬럼(G,H열)로 강제 지정
    ota_price_cols = [df_ota.columns[6], df_ota.columns[7]]

result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
wb = load_workbook(result_path)
ws = wb.active
# 비교로그 시트 생성(기존 있으면 삭제)
if '비교로그' in wb.sheetnames:
    del wb['비교로그']
log_ws = wb.create_sheet('비교로그')
log_ws.append(['고객명', '전체매출 행번호', '전체매출 가격', '파일명', '행번호', '비교 가격', '원가격'])

# Remittances 파일별 행 오프셋 기록비교 가격
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

# 1. 전체고객목록에서 아고다인 고객명별로 인덱스와 가격(합계, 객실료) 수집
grouped_rows = defaultdict(list)
for idx, row in df_all.iterrows():
    vendor = str(row.get('거래처', '')).strip()
    if vendor != '아고다':
        continue
    name = str(row.get(col_name_all, '')).strip()
    price1 = str(row.get(col_price_all_1, '')).replace(',', '').strip()
    price2 = str(row.get(col_price_all_2, '')).replace(',', '').strip()
    try:
        price1_f = float(price1)
    except:
        price1_f = 0.0
    try:
        price2_f = float(price2)
    except:
        price2_f = 0.0
    # 합계 우선, 없으면 객실료
    use_price = price2_f if price2_f else price1_f
    grouped_rows[name].append((idx, use_price))

# 2. Remittances에서 이름별 금액 리스트 수집
otas_by_name = defaultdict(list)
for _, row in df_ota.iterrows():
    name = str(row[col_name_ota]).strip()
    for price_col in ota_price_cols:
        try:
            price = float(str(row[price_col]).replace(',', '').strip())
            otas_by_name[name].append(price)
        except:
            continue

# 3. 매칭 처리
used_ota_idx = set()
for name, rows in grouped_rows.items():
    total_price = sum(price for _, price in rows)
    # Remittances에서 해당 이름의 금액 중 합과 일치하는 것 찾기
    found = False
    for i, price in enumerate(otas_by_name.get(name, [])):
        if price == total_price and (name, i) not in used_ota_idx:
            # 매칭된 Remittances 금액은 중복 사용 방지
            used_ota_idx.add((name, i))
            found = True
            break
    if found:
        # 전체고객목록의 해당 이름 모든 행을 노란색으로 표시
        for idx, _ in rows:
            for cell in ws[idx+2]:
                cell.fill = fill_yellow
    else:
        # 기존 개별 비교 로직(잔여 케이스) 수행
        for idx, _ in rows:
            row = df_all.iloc[idx]
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
            # 합계 우선, 없으면 객실료
            use_price = price2_f if price2_f else price1_f
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
                                    log_info = [name, idx+2, use_price, fname, file_row, str(match_row[price_col]), str(match_row[price_col])]
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
                    else:
                        # log_info가 None인 경우에도 비교로그에 한 줄 남김
                        log_ws.append([name, idx+2, use_price, '-', '-', '불일치', '-'])
    # Remittances에 없는 고객명은 아무 표시도 하지 않음


# 결과 저장
result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
wb.save(result_path)
print(f'완료: {result_path}에 저장됨')

# Peter Ludwig 비교로그 출력
print_peter_ludwig_log()
