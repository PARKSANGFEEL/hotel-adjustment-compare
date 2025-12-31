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

# 파일 경로 설정
dir_base = os.path.dirname(os.path.abspath(__file__))

# (Booking.com PDF 비교 코드 완전 삭제)

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

# 파일 경로 설정
dir_base = os.path.dirname(os.path.abspath(__file__))


# (Booking.com PDF 비교 코드 완전 삭제)
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

# 파일 경로 설정
dir_base = os.path.dirname(os.path.abspath(__file__))

# (Booking.com PDF 비교 코드 완전 삭제)

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

# 파일 경로 설정
dir_base = os.path.dirname(os.path.abspath(__file__))


# (Booking.com PDF 비교 코드 완전 삭제)

# (Booking.com PDF 비교 코드 완전 삭제)
