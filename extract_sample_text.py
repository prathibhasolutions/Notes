from PyPDF2 import PdfReader

pdf_path = 'Physics arihant handbook.pdf'
output_path = 'sample_text.txt'

reader = PdfReader(pdf_path)
text = ''
for i in range(min(5, len(reader.pages))):
    text += reader.pages[i].extract_text() or ''

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(text)
print('Sample text written to sample_text.txt')
