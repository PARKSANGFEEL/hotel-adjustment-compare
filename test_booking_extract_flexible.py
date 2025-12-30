import PyPDF2
import re

pdf_path = 'ota-adjustment/부킹11월.pdf'

with open(pdf_path, 'rb') as f:
    reader = PyPDF2.PdfReader(f)
    text = ''
    for page in reader.pages:
        text += page.extract_text() + '\n'

lines = [line.strip() for line in text.split('\n') if line.strip()]

results = []
i = 0
while i < len(lines):
    if lines[i] == '예약' and i+1 < len(lines) and re.match(r'\d{10}', lines[i+1]):
        ref = lines[i+1]
        # 이름, 금액, KRW 찾기 (10줄 이내)
        name, amount = None, None
        for j in range(i+2, min(i+15, len(lines))):
            if re.match(r'^[A-Z가-힣][A-Za-z가-힣 ]+$', lines[j]):
                name = lines[j]
            if lines[j] == 'KRW' and j+1 < len(lines):
                amount = lines[j+1].replace(',', '')
                break
        if ref and name and amount:
            results.append((ref, name, amount))
        i = j+2 if amount else i+2
    else:
        i += 1
for ref, name, amount in results:
    print('참조번호:', ref, '이름:', name, '금액:', amount)
