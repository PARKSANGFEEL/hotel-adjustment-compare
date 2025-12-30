import PyPDF2
pdf_path = 'ota-adjustment/부킹11월.pdf'
with open(pdf_path, 'rb') as f:
    reader = PyPDF2.PdfReader(f)
    text = ''
    for page in reader.pages:
        text += page.extract_text() + '\n'
with open('booking_pdf_alltext.txt', 'w', encoding='utf-8') as out:
    out.write(text)
print('saved')
