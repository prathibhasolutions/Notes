from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib import colors


def build_excel_notes(output_path: str) -> None:
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2.2 * cm,
        rightMargin=2.2 * cm,
        topMargin=2.2 * cm,
        bottomMargin=2.2 * cm,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "Title",
        parent=styles["Heading1"],
        fontSize=20,
        leading=26,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1F4E78"),
        spaceAfter=12,
    )

    subtitle_style = ParagraphStyle(
        "Subtitle",
        parent=styles["Normal"],
        fontSize=12,
        leading=16,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#2E75B5"),
        spaceAfter=12,
    )

    h1 = ParagraphStyle(
        "H1",
        parent=styles["Heading2"],
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#2E75B5"),
        spaceAfter=6,
        spaceBefore=8,
    )

    h2 = ParagraphStyle(
        "H2",
        parent=styles["Heading3"],
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#4472C4"),
        spaceAfter=4,
        spaceBefore=6,
    )

    body = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=10.5,
        leading=14.5,
        alignment=TA_LEFT,
        spaceAfter=4,
    )

    bullet = ParagraphStyle(
        "Bullet",
        parent=styles["Normal"],
        fontSize=10.5,
        leading=14.5,
        leftIndent=12,
        bulletIndent=6,
        spaceAfter=2,
    )

    today = date(2026, 2, 6).strftime("%d-%m-%Y")

    story = []

    story.append(Paragraph("Prathibha Institute", title_style))
    story.append(Paragraph("MS Excel Notes (Beginner Level)", subtitle_style))
    story.append(Paragraph(f"Prepared for print | Date: {today}", subtitle_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "These notes are original learning material created for beginners. They follow the syllabus provided and include exercises for practice.",
        body,
    ))
    story.append(PageBreak())

    # 1. Introduction
    story.append(Paragraph("1. Excel Introduction", h1))
    story.append(Paragraph(
        "Microsoft Excel is a spreadsheet application used to store data in rows and columns, perform calculations, and create charts. It is widely used in offices, education, finance, and data analysis.",
        body,
    ))
    story.append(Paragraph("Key uses:", h2))
    for item in [
        "Data entry and organization",
        "Calculations using formulas",
        "Data analysis with sorting and filtering",
        "Visualization using charts",
        "Reporting and printing",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 2. Get Started / Overview
    story.append(Paragraph("2. Get Started and Overview", h1))
    story.append(Paragraph(
        "A workbook contains one or more worksheets. Each worksheet has rows (1, 2, 3...) and columns (A, B, C...). A cell is identified by its column and row (e.g., A1).",
        body,
    ))
    story.append(Paragraph("Excel interface basics:", h2))
    for item in [
        "Ribbon: Tabs with tools (Home, Insert, Page Layout, Formulas, Data, Review, View)",
        "Name Box: Shows selected cell address",
        "Formula Bar: View and edit cell contents",
        "Sheet Tabs: Switch between worksheets",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 3. Syntax, Cells, Ranges
    story.append(Paragraph("3. Syntax, Cells, and Ranges", h1))
    story.append(Paragraph(
        "Formulas always start with =. A formula can include numbers, cell references, operators, and functions.",
        body,
    ))
    story.append(Paragraph("Examples:", h2))
    for item in [
        "=A1+B1 (adds two cells)",
        "=SUM(A1:A10) (adds a range)",
        "=A1*10 (multiplies a cell by 10)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))
    story.append(Paragraph("Cells and ranges:", h2))
    for item in [
        "Single cell: B3",
        "Range: B3:D10",
        "Entire column: B:B",
        "Entire row: 3:3",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 4. Fill, Move, Add, Delete, Undo
    story.append(Paragraph("4. Fill, Move, Add, Delete, Undo", h1))
    story.append(Paragraph(
        "Use the fill handle (small square at the bottom-right of a selected cell) to copy or extend patterns. You can move or copy cells via drag-and-drop or cut/copy/paste. Insert or delete rows/columns from the Home tab.",
        body,
    ))
    story.append(Paragraph("Quick tips:", h2))
    for item in [
        "Ctrl+Z = Undo, Ctrl+Y = Redo",
        "Ctrl+C = Copy, Ctrl+X = Cut, Ctrl+V = Paste",
        "Right-click row/column header to insert or delete",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 5. Formulas and References
    story.append(Paragraph("5. Formulas and References", h1))
    story.append(Paragraph(
        "Excel supports arithmetic operators: +, -, *, /, ^, and parentheses for order of operations.",
        body,
    ))
    story.append(Paragraph("Reference types:", h2))
    for item in [
        "Relative: A1 (changes when copied)",
        "Absolute: $A$1 (fixed when copied)",
        "Mixed: $A1 or A$1 (row or column fixed)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 6. Common Functions
    story.append(Paragraph("6. Common Functions", h1))
    story.append(Paragraph(
        "Functions are built-in formulas. Syntax example: =FUNCTION(arguments)",
        body,
    ))

    story.append(Paragraph("Logical:", h2))
    for item in [
        "=IF(test, value_if_true, value_if_false)",
        "=AND(condition1, condition2)",
        "=OR(condition1, condition2)",
        "=XOR(condition1, condition2)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    story.append(Paragraph("Math & Statistics:", h2))
    for item in [
        "=SUM(range)",
        "=AVERAGE(range)",
        "=MAX(range)",
        "=MIN(range)",
        "=COUNT(range)",
        "=COUNTIF(range, criteria)",
        "=COUNTIFS(range1, criteria1, range2, criteria2)",
        "=SUMIF(range, criteria, sum_range)",
        "=SUMIFS(sum_range, range1, criteria1, ...)",
        "=RAND()",
        "=NPV(rate, values)",
        "=MEDIAN(range)",
        "=MODE(range)",
        "=STDEV.P(range)",
        "=STDEV.S(range)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    story.append(Paragraph("Text:", h2))
    for item in [
        "=CONCAT(text1, text2, ...)",
        "=LEFT(text, num_chars)",
        "=RIGHT(text, num_chars)",
        "=LOWER(text)",
        "=TRIM(text)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    story.append(Paragraph("Date:", h2))
    for item in [
        "=TODAY()",
        "=DATE(year, month, day)",
        "=YEAR(date)",
        "=MONTH(date)",
        "=DAY(date)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    story.append(Paragraph("Lookup:", h2))
    for item in [
        "=VLOOKUP(lookup_value, table_array, col_index, [range_lookup])",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 7. Formatting
    story.append(Paragraph("7. Formatting", h1))
    story.append(Paragraph(
        "Formatting improves readability. Use font styles, colors, borders, number formats, and alignment to present data clearly.",
        body,
    ))
    story.append(Paragraph("Common formatting tools:", h2))
    for item in [
        "Font name, size, bold/italic/underline",
        "Cell fill color and font color",
        "Borders and gridlines",
        "Number formats (General, Number, Currency, Percentage, Date)",
        "Format Painter to copy formatting",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 8. Tables
    story.append(Paragraph("8. Tables", h1))
    story.append(Paragraph(
        "Convert a data range into a table to enable filtering, structured references, and consistent formatting.",
        body,
    ))
    story.append(Paragraph("Table benefits:", h2))
    for item in [
        "Automatic headers and filters",
        "Alternating row colors",
        "Structured references (e.g., Table1[Amount])",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 9. Data Analysis: Sort & Filter
    story.append(Paragraph("9. Data Analysis: Sort and Filter", h1))
    story.append(Paragraph(
        "Sorting arranges data in ascending or descending order. Filtering shows only rows that meet criteria.",
        body,
    ))
    story.append(Paragraph("Sorting examples:", h2))
    for item in [
        "Sort by Name A–Z",
        "Sort by Date (Newest to Oldest)",
        "Sort by multiple columns (Department then Name)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    story.append(Paragraph("Filtering examples:", h2))
    for item in [
        "Filter Sales > 5000",
        "Filter Category = "
        "Electronics"
        "",
        "Filter by month in a Date column",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 10. Conditional Formatting
    story.append(Paragraph("10. Conditional Formatting", h1))
    story.append(Paragraph(
        "Conditional formatting highlights data based on rules, helping you spot trends and exceptions.",
        body,
    ))
    story.append(Paragraph("Common rules:", h2))
    for item in [
        "Highlight Cell Rules (greater than, less than, between)",
        "Top/Bottom rules (top 10, bottom 10)",
        "Data Bars, Color Scales, Icon Sets",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 11. Charts
    story.append(Paragraph("11. Charts", h1))
    story.append(Paragraph(
        "Charts visualize data. Choose a chart type based on what you want to show: comparisons, trends, or parts of a whole.",
        body,
    ))
    story.append(Paragraph("Common chart types:", h2))
    for item in [
        "Column/Bar: compare categories",
        "Line: show trends over time",
        "Pie: show proportions",
        "Scatter: show relationships between two variables",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 12. Pivot Tables
    story.append(Paragraph("12. Pivot Tables", h1))
    story.append(Paragraph(
        "Pivot tables summarize large datasets quickly by grouping and aggregating fields.",
        body,
    ))
    story.append(Paragraph("Basic steps:", h2))
    for item in [
        "Select your data range and choose Insert > PivotTable",
        "Drag fields to Rows, Columns, Values, and Filters",
        "Change aggregation (Sum, Count, Average)",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # 13. Page Layout and Printing
    story.append(Paragraph("13. Page Layout and Printing", h1))
    story.append(Paragraph(
        "Use Page Layout to control margins, orientation, size, headers/footers, and print area for professional output.",
        body,
    ))
    story.append(Paragraph("Printing checklist:", h2))
    for item in [
        "Set print area",
        "Fit sheet on one page if needed",
        "Use Print Preview",
        "Add headers/footers with page numbers",
    ]:
        story.append(Paragraph(f"• {item}", bullet))

    # Exercises
    story.append(PageBreak())
    story.append(Paragraph("Practice Exercises", h1))
    story.append(Paragraph(
        "Create a new workbook and complete the following exercises. Save your file as Excel_Practice.xlsx.",
        body,
    ))

    exercises = [
        "Enter a table of 10 students with columns: Roll No, Name, Class, English, Maths, Science. Calculate Total and Average using formulas.",
        "Apply formatting: bold headers, borders, and a light fill color for the header row.",
        "Use conditional formatting to highlight any subject mark below 35 in red.",
        "Sort the student list by Total (highest to lowest).",
        "Filter students who scored 60 or more in Maths.",
        "Create a column chart of student totals.",
        "Create a table from the student data and rename it as StudentsTable.",
        "Use COUNTIF to count how many students passed Maths (>= 35).",
        "Use AVERAGEIF to find the average Maths score for students who passed.",
        "Create a sales table with 12 rows (one per month) and columns: Month, Product, Units, Price. Calculate Revenue = Units*Price.",
        "Use SUMIFS to compute total revenue for a selected product.",
        "Create a pivot table summarizing total revenue by Product.",
        "Use VLOOKUP to fetch Price for a product from a separate price list table.",
        "Set a print area and print preview the student report on a single page.",
    ]

    for idx, task in enumerate(exercises, start=1):
        story.append(Paragraph(f"{idx}. {task}", body))

    doc.build(story)


if __name__ == "__main__":
    output_pdf = "Prathibha_Institute_MS_Excel_Notes.pdf"
    build_excel_notes(output_pdf)
    print(f"Generated: {output_pdf}")
