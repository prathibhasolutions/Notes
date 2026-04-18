from PyPDF2 import PdfReader

files = [
    "01 - Matrices.pdf",
    "01 - Vector Algebra.pdf",
    "02 - Determinants.pdf",
    "02 - Three Dimensional Geometry.pdf",
    "05 - Straight Lines.pdf",
]

for f in files:
    reader = PdfReader(f)
    text = ""
    for i in range(min(3, len(reader.pages))):
        text += (reader.pages[i].extract_text() or "") + "\n"

    print(f"\n===== {f} =====")
    print(text[:4000].replace("\r", " "))
    print(f"--- pages: {len(reader.pages)}")
