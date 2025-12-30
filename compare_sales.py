
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
import tabula
import re

# 파일 경로 설정
dir_base = os.path.dirname(os.path.abspath(__file__))

directory_ota = os.path.join(dir_base, 'ota-adjustment')

# 결과 엑셀 파일에서만 데이터 로드
result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
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


# tabula-py 기반 Booking.com PDF 표 추출 함수
def extract_booking_pdf_data(pdf_path):
    # Sonia Mushime Simbi의 원본금액에서 564570이 나오도록 곱해야 하는 비율 계산용
    target_name = 'Sonia Mushime Simbi'
    target_result = 564570
    found_ratio = None
    # Booking.com PDF는 stream=True, lattice=False가 가장 잘 맞음
    data = []
    try:
        tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True, stream=True, lattice=False)
        if not tables or len(tables) == 0:
            print(f"[진단] tabula stream=True로 표 추출 실패")
            return data
        # 표 개수 및 미리보기 로그 제거
        # 모든 표에서 참조번호 추출
        for i, df in enumerate(tables):
            try:
                df.columns = ['유형', '참조번호', '유형2', '체크인', '체크아웃', '투숙객이름', '통화', '금액']
            except Exception:
                continue
            for _, row in df.iterrows():
                try:
                    ref = str(row['참조번호']).strip()
                    name = str(row['투숙객이름']).strip()
                    amount = str(row['금액']).replace(',', '').strip()
                    if name == target_name:
                        print(f"[진단-부킹PDF] Sonia Mushime Simbi: ref={ref}, 원본금액={amount}")
                        try:
                            amt_f = float(amount)
                            ratio = target_result / amt_f
                            found_ratio = ratio
                            print(f"[진단-부킹PDF] Sonia Mushime Simbi: {amt_f}에 곱해야 564570이 되는 비율={ratio}")
                        except Exception:
                            pass
                    if re.match(r'^\d{10}$', ref) and name and amount.replace('.', '', 1).isdigit():
                        data.append({'ref': ref, 'name': name, 'amount': float(amount)})
                except Exception:
                    continue
    except Exception as e:
        print(f"[진단] tabula 예외: {e}")
    return data
# 부킹으로 시작되는 PDF 파일 처리
booking_files = [f for f in os.listdir(directory_ota) if f.startswith('부킹') and f.endswith('.pdf')]
booking_data = []
for file in booking_files:
    path = os.path.join(directory_ota, file)
    booking_data.extend(extract_booking_pdf_data(path))

# booking_data 추출 결과 로그 제거

# Booking.com 매출 비교 및 색상 처리
def compare_booking():
    # KAYO HIRAI 진단용 로그
    # PDF booking_data에서 KAYO HIRAI 추출
    kayo_pdf_amount = None
    for d in booking_data:
        if normalize(d.get('name','')) == normalize('KAYO HIRAI'):
            kayo_pdf_amount = d['amount']
            print(f"[진단-KAYO HIRAI] PDF 추출: name='{d.get('name','')}', normalize='{normalize(d.get('name',''))}', amount={d['amount']}")
            break
    # 엑셀 73행(헤더 포함, 실제 데이터는 73행)
    kayo_excel_name = ws.cell(row=73, column=4).value
    kayo_excel_k = ws.cell(row=73, column=11).value
    print(f"[진단-KAYO HIRAI] 엑셀 73행: name='{kayo_excel_name}', normalize='{normalize(str(kayo_excel_name))}', K열='{kayo_excel_k}'")
    try:
        kayo_excel_k_f = float(str(kayo_excel_k).replace(",", "").replace(" ", "").strip())
    except:
        kayo_excel_k_f = None
    if kayo_pdf_amount is not None and kayo_excel_k_f is not None:
        print(f"[진단-KAYO HIRAI] 비교: 엑셀 K열={kayo_excel_k_f}, PDF×0.82={round(kayo_pdf_amount*0.82)}, 일치여부={'일치' if round(kayo_excel_k_f)==round(kayo_pdf_amount*0.82) else '불일치'}")
    processed_idx = set()
    # PDF booking_data의 모든 고객명에 대해 자동 처리
    pdf_names = set(normalize(d['name']) for d in booking_data if d.get('name'))
    for pdf_name in pdf_names:
        # 시트에서 해당 고객명의 모든 행 찾기
        row_list = []
        k_sum = 0
        for row_idx in range(2, ws.max_row+1):
            name_cell = ws.cell(row=row_idx, column=4).value
            if normalize(str(name_cell)) == pdf_name:
                row_list.append(row_idx)
                val = ws.cell(row=row_idx, column=11).value
                print(f"[진단-{pdf_name}] 행:{row_idx}, K열:'{val}'")
                try:
                    k_sum += float(str(val).replace(",", "").replace(" ", "").strip())
                except:
                    pass
        # PDF에서 해당 고객명 금액 추출
        pdf_amount = None
        for d in booking_data:
            if normalize(d['name']) == pdf_name:
                pdf_amount = d['amount']
                break
        print(f"[진단-{pdf_name}] 전체 K열 합계: {k_sum}, PDF 원본금액: {pdf_amount}, PDF×0.82: {round(pdf_amount*0.82) if pdf_amount else None}, 비교결과: {'일치' if (pdf_amount is not None and round(k_sum) == round(pdf_amount*0.82)) else '불일치'} (행:{row_list})")
        if pdf_amount is not None and round(k_sum) == round(pdf_amount * 0.82):
            for r in row_list:
                for cell in ws[r]:
                    cell.fill = fill_yellow
                    cell.font = Font(color="000000")
            for r in row_list:
                processed_idx.add(r-2)
        # 합산이 일치하지 않으면 processed_idx에 추가하지 않음 (빨간색+비교로그 기록을 위해)
    for idx, row in df_all.iterrows():
        if idx in processed_idx:
            continue
        ws_row = idx + 2
        vendor = str(row.get('거래처', '')).strip()
        if vendor != '부킹닷컴':
            continue
        ota_no = str(row.get('OTA번호', '')).strip()[:10]
        price1 = str(row.get('객실료', '')).replace(',', '').strip()
        price2 = str(row.get('합계', '')).replace(',', '').strip()
        try:
            price1_f = float(price1)
        except:
            price1_f = None
        try:
            price2_f = float(price2)
        except:
            price2_f = None
        match = None
        for d in booking_data:
            if str(d['ref']).strip() == str(ota_no).strip():
                match = d
                break
        if not match:
            print(f"[진단-부킹] 매칭 실패: idx={ws_row}, OTA번호={ota_no}, 고객명={row.get('고객명','')}")
        if match:
            pdf_amount = round(match['amount'] * 0.82)
            pdf_original = match['amount']
            if (price1_f is not None and pdf_amount == price1_f) or (price2_f is not None and pdf_amount == price2_f):
                for cell in ws[ws_row]:
                    cell.fill = fill_yellow
                processed_idx.add(idx)
            else:
                sum_rows = []
                sum_total = 0
                기준_ota = normalize(ota_no)
                기준_name = normalize(row.get('고객명',''))
                for j in range(idx, ws.max_row-1):
                    next_ota = normalize(ws.cell(row=j+2, column=3).value)
                    next_name = normalize(ws.cell(row=j+2, column=4).value)
                    if next_ota == 기준_ota and next_name == 기준_name:
                        val = ws.cell(row=j+2, column=7).value
                        if val is None:
                            val = ws.cell(row=j+2, column=6).value
                        try:
                            val_f = float(str(val).replace(",", "").replace(" ", "").strip())
                        except:
                            val_f = 0
                        sum_total += val_f
                        sum_rows.append(j+2)
                    else:
                        break
                if len(sum_rows) > 1 and round(sum_total) == pdf_amount:
                    for r in sum_rows:
                        for cell in ws[r]:
                            cell.fill = fill_yellow
                            cell.font = Font(color="000000")
                    for r in sum_rows:
                        ws_name = normalize(ws.cell(row=r, column=4).value)
                        ws_ota = normalize(ws.cell(row=r, column=3).value)
                        for idx2, row2 in df_all.iterrows():
                            row_name = normalize(row2.get('고객명',''))
                            row_ota = normalize(row2.get('OTA번호',''))
                            if row_name == ws_name and row_ota == ws_ota:
                                processed_idx.add(idx2)
                    continue
                if idx not in processed_idx:
                    print(f"[진단-부킹] 금액 불일치: idx={ws_row}, OTA번호={ota_no}, 고객명={row.get('고객명','')}, PDF원본금액={pdf_original}, 비교금액(×0.82)={pdf_amount}, 객실료={price1_f}, 합계={price2_f}")
                    for cell in ws[ws_row]:
                        cell.font = font_red
                    log_ws.append([row.get('고객명', ''), ws_row, price1, '부킹PDF', '-', str(pdf_amount)])


compare_booking()
# 결과 저장
result_path = os.path.join(dir_base, '매출_검토_결과.xlsx')
wb.save(result_path)
print(f'완료: {result_path}에 저장됨')

# Peter Ludwig 비교로그 출력
print_peter_ludwig_log()
