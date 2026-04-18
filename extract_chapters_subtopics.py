import os
from pdfminer.high_level import extract_text
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import re

PDF_PATH = 'Physics arihant handbook.pdf'
OUTPUT_PATH = 'Physics_Arihant_Handbook_Chapters_Subtopics.pdf'

# Step 1: Extract text from PDF

def extract_chapters_and_subtopics(pdf_path):
    from PyPDF2 import PdfReader
    reader = PdfReader(pdf_path)
    text = ''
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + '\n'
    chapters = []
    current_chapter = None
    # Match chapter lines like '1. Units and Measurement 1-8' or 'Chapter 1: ...'
    chapter_pattern = re.compile(r'^(\d+\.\s*[^\d]+\d+-\d+|Chapter\s*\d+\s*[:.-]?\s*.+)', re.IGNORECASE)
    # Match subtopics starting with dash or bullet
    subtopic_pattern = re.compile(r'^[—\-•]\s*(.+)')
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        chapter_match = chapter_pattern.match(line)
        if chapter_match:
            if current_chapter:
                chapters.append(current_chapter)
            current_chapter = {'name': chapter_match.group(0), 'subtopics': []}
        elif current_chapter and subtopic_pattern.match(line):
            current_chapter['subtopics'].append(subtopic_pattern.match(line).group(1))
    if current_chapter:
        chapters.append(current_chapter)
    return chapters

# Step 2: Generate new PDF

def generate_pdf(chapters, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    y = height - 50
    c.setFont('Helvetica-Bold', 16)
    c.drawString(50, y, 'Physics Arihant Handbook: Chapters & Subtopics')
    y -= 40
    c.setFont('Helvetica', 12)
    for chapter in chapters:
        if y < 100:
            c.showPage()
            y = height - 50
        c.setFont('Helvetica-Bold', 14)
        c.drawString(50, y, f"Chapter: {chapter['name']}")
        y -= 25
        c.setFont('Helvetica', 12)
        for subtopic in chapter['subtopics']:
            if y < 60:
                c.showPage()
                y = height - 50
            c.drawString(70, y, f"- {subtopic}")
            y -= 18
        y -= 10
    c.save()

if __name__ == '__main__':
    chapters = extract_chapters_and_subtopics(PDF_PATH)
    generate_pdf(chapters, OUTPUT_PATH)
    print(f"Generated: {OUTPUT_PATH}")
