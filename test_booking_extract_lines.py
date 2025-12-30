import PyPDF2
import re

pdf_path = 'ota-adjustment/부킹11월.pdf'

with open(pdf_path, 'rb') as f:
    reader = PyPDF2.PdfReader(f)
    text = ''
    for page in reader.pages:
        text += page.extract_text() + '\n'

lines = [line.strip() for line in text.split('\n') if line.strip()]

# 슬라이딩 윈도우로 10줄씩 묶어서 패턴 찾기
def find_booking_entries(lines):
    results = []
    for i in range(len(lines) - 7):
        # 패턴: 예약, 참조번호(숫자), 예약, ... 투숙객이름, KRW, 금액
        if (lines[i] == '예약' and
            re.match(r'\d{10}', lines[i+1]) and
            lines[i+2] == '예약' and
            'KRW' in lines[i+6] and
            re.match(r'[\d,]+\.?\d*', lines[i+7])):
            ref = lines[i+1]
            name = lines[i+5]
            amount = lines[i+7].replace(',', '')
            results.append((ref, name, amount))
    return results

for ref, name, amount in find_booking_entries(lines):
    print('참조번호:', ref, '이름:', name, '금액:', amount)
