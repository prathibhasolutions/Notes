import json
import random
from pathlib import Path

# Set seed for reproducibility (optional)
random.seed(42)

input_dir = Path("extracted_questions_json")
output_file = Path("maths_file.json")

pdf_files = [
    "01___Matrices.json",
    "01___Vector_Algebra.json",
    "02___Determinants.json",
    "02___Three_Dimensional_Geometry.json",
    "05___Straight_Lines.json",
]

combined_questions = []

for pdf_file in pdf_files:
    file_path = input_dir / pdf_file
    with open(file_path, "r", encoding="utf-8") as f:
        questions = json.load(f)
    
    # Randomly sample 30 questions from each file
    sampled = random.sample(questions, min(30, len(questions)))
    combined_questions.extend(sampled)
    
    print(f"{pdf_file}: {len(sampled)} questions sampled (total available: {len(questions)})")

# Shuffle all combined questions for better mixing
random.shuffle(combined_questions)

# Write to output file
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(combined_questions, f, ensure_ascii=False, indent=2)

print(f"\nTotal questions: {len(combined_questions)}")
print(f"Output file: {output_file}")
