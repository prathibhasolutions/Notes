import pdfplumber
import re
from pathlib import Path
from jinja2 import Template

# Path to the Chemistry Arihant Handbook PDF
PDF_PATH = "Chemistry arihant handbook.pdf"
OUTPUT_HTML = "alcohols_phenols_ethers.html"

# HTML template (simplified, you can expand as needed)
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{{ chapter_title }}</title>
    <link rel="stylesheet" href="../styles.css">
    <link rel="icon" type="image/png" href="../../../logo.png">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <script type="text/javascript" async src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.7/MathJax.js?config=TeX-MML-AM_CHTML"></script>
</head>
<body class="bg-light">
<section class="header" style="font-family: sans-serif;color: rgb(23, 41, 84);">
    <nav>
        <img src="../logo.png" alt="" class="logo">
    </nav>
</section>
<br><br><br><br>
<div class="container mt-5">
    <h1 class="text-center mb-4">{{ chapter_title }}</h1>
    <div class="card shadow-lg">
        <div class="card-header bg-primary text-white">
            <h4>{{ chapter_title }}</h4>
        </div>
        <div class="card-body">
            {% for para in content %}
                <p>{{ para }}</p>
            {% endfor %}
        </div>
    </div>
</div>
<footer id="footer" class="footer" style="background-color: rgb(23, 41, 84); color: #ffffff; padding: 10px 0;">
    <div class="container text-center">
        <p>&copy; 2024 Prathibha Institute. All rights reserved.</p>
        <p>Designed by <a href="https://www.prathibhasolutions.com" style="color:aquamarine;">Prathibha Solutions</a></p>
    </div>
</footer>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
'''

def extract_chapter(pdf_path, chapter_name):
    """
    Extracts the text of a chapter from the PDF by chapter name.
    Returns a list of paragraphs.
    """
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    # Find chapter start and end
    pattern = re.compile(rf"{chapter_name}.*?(?=\n[A-Z][A-Za-z\s]+\n|\Z)", re.DOTALL)
    match = pattern.search(text)
    if match:
        chapter_text = match.group(0)
        # Split into paragraphs
        paragraphs = [p.strip() for p in chapter_text.split('\n') if p.strip()]
        return paragraphs
    else:
        return ["Chapter not found or extraction failed."]

def main():
    chapter_name = "Alcohols, Phenols and Ethers"
    content = extract_chapter(PDF_PATH, chapter_name)
    template = Template(HTML_TEMPLATE)
    html = template.render(chapter_title=chapter_name, content=content)
    Path(OUTPUT_HTML).write_text(html, encoding="utf-8")
    print(f"HTML page generated: {OUTPUT_HTML}")

if __name__ == "__main__":
    main()
