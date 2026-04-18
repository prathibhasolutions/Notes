from PyPDF2 import PdfReader

pdf_path = 'Physics arihant handbook.pdf'
output_path = 'sample_text_6_20.txt'

reader = PdfReader(pdf_path)
text = ''
for i in range(5, min(20, len(reader.pages))):
    text += reader.pages[i].extract_text() or ''

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(text)
print('Sample text written to sample_text_6_20.txt')
