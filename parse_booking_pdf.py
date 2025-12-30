import PyPDF2
import re

pdf_path = 'ota-adjustment/부킹11월.pdf'

with open(pdf_path, 'rb') as f:
    reader = PyPDF2.PdfReader(f)
    text = ''
    for page in reader.pages:
        text += page.extract_text() + '\n'

# 참조번호(10자리), 투숙객이름, 금액 추출 예시 (패턴은 실제 PDF 구조에 맞게 조정 필요)
# 예: 참조번호: 1234567890, 투숙객: 홍길동, 금액: 123,456
ref_pattern = r'(\d{10})'
name_pattern = r'([가-힣A-Za-z ]{2,})'
amount_pattern = r'([\d,]+)'

# 예시: 한 줄에 참조번호, 이름, 금액이 모두 있는 경우
pattern = re.compile(rf'{ref_pattern}.*?{name_pattern}.*?{amount_pattern}')

matches = pattern.findall(text)
for m in matches:
    print('참조번호:', m[0], '이름:', m[1], '금액:', m[2])
