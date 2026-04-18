import json
import re
from pathlib import Path
from PyPDF2 import PdfReader
import pdfplumber

PDF_FILES = [
    "01 - Matrices.pdf",
    "01 - Vector Algebra.pdf",
    "02 - Determinants.pdf",
    "02 - Three Dimensional Geometry.pdf",
    "05 - Straight Lines.pdf",
]

Q_SPLIT = re.compile(r"(?=\bQ\d+\.)")
Q_START = re.compile(r"^Q\d+\.")
OPT_PAT = re.compile(r"\((\d)\)\s*(.*?)(?=\s*\(\d\)|$)", re.DOTALL)


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def read_pdf_text(path: Path) -> str:
    chunks = []

    # pdfplumber generally preserves ordering and symbols better for these papers.
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                if t:
                    chunks.append(t)
    except Exception:
        chunks = []

    if chunks:
        return "\n".join(chunks)

    # Fallback parser for problematic files/pages.
    reader = PdfReader(str(path))
    for page in reader.pages:
        try:
            t = page.extract_text() or ""
            if t:
                chunks.append(t)
        except Exception:
            continue
    return "\n".join(chunks)


def parse_questions(raw_text: str):
    parts = Q_SPLIT.split(raw_text)
    questions = []

    for part in parts:
        p = part.strip()
        if not p or not Q_START.match(p):
            continue

        # Stop before boilerplate if present
        p = p.split("For solutions, download the MARKS App")[0].strip()

        options = OPT_PAT.findall(p)
        q_text = p
        if options:
            first_opt = re.search(r"\(1\)", p)
            if first_opt:
                q_text = p[: first_opt.start()].strip()

        # Remove leading Qn.
        q_text = re.sub(r"^Q\d+\.\s*", "", q_text).strip()

        if not q_text:
            continue

        option_items = []
        if options:
            for _, opt_text in options[:4]:
                option_items.append(
                    {
                        "option_text": normalize_text(opt_text),
                        "is_correct": False,
                    }
                )

        # Ensure schema shape with 4 options when present in source
        if option_items and len(option_items) < 4:
            for _ in range(4 - len(option_items)):
                option_items.append({"option_text": "", "is_correct": False})

        questions.append(
            {
                "question_text": normalize_text(q_text),
                "solution_text": "",
                "options": option_items,
            }
        )

    return questions


def main():
    base = Path(__file__).resolve().parent
    out_dir = base / "extracted_questions_json"
    out_dir.mkdir(exist_ok=True)

    summary = {}

    for pdf in PDF_FILES:
        pdf_path = base / pdf
        raw_text = read_pdf_text(pdf_path)
        questions = parse_questions(raw_text)

        out_name = pdf_path.stem.replace(" ", "_").replace("-", "_") + ".json"
        out_path = out_dir / out_name
        with out_path.open("w", encoding="utf-8") as f:
            json.dump(questions, f, ensure_ascii=False, indent=2)

        summary[pdf] = {
            "count": len(questions),
            "output": str(out_path.name),
        }

    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
