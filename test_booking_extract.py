import PyPDF2
import re

pdf_path = 'ota-adjustment/부킹11월.pdf'

with open(pdf_path, 'rb') as f:
    reader = PyPDF2.PdfReader(f)
    text = ''
    for page in reader.pages:
        text += page.extract_text() + '\n'

# JUN HAN WEE만 추출
pattern = re.compile(r'(\d{10}).*?(JUN HAN WEE).*?([\d,]+\.?\d*)')
for m in pattern.findall(text):
    print('참조번호:', m[0], '이름:', m[1], '금액:', m[2])
