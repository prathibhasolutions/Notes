from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib import colors


def build_request_letter(output_path: str) -> None:
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2.2 * cm,
        rightMargin=2.2 * cm,
        topMargin=2.5 * cm,
        bottomMargin=2.5 * cm,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "Title",
        parent=styles["Heading2"],
        fontSize=14,
        leading=18,
        alignment=TA_LEFT,
        textColor=colors.black,
        spaceAfter=10,
    )

    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=11,
        leading=16,
        alignment=TA_LEFT,
        spaceAfter=8,
    )

    small_style = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontSize=10,
        leading=14,
        alignment=TA_LEFT,
        spaceAfter=6,
    )

    today = date(2026, 2, 24).strftime("%d-%m-%Y")

    story = []

    story.append(Paragraph("Request for Sanitary Certificate", title_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph(f"Date: {today}", small_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("From:", small_style))
    story.append(Paragraph("President", body_style))
    story.append(Paragraph("Prathibha Educational Society", body_style))
    story.append(Paragraph("6-12-63, Namdevwada", body_style))
    story.append(Paragraph("Nizamabad-503002", body_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("To:", small_style))
    story.append(Paragraph("The Municipal Commissioner", body_style))
    story.append(Paragraph("Nizamabad Municipal Corporation", body_style))
    story.append(Paragraph("Nizamabad", body_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Subject: Request for issue of Sanitary Certificate", body_style))
    story.append(Spacer(1, 4))

    story.append(Paragraph("Respected Sir/Madam,", body_style))

    body_text = (
        "I, the undersigned, on behalf of Prathibha Educational Society, request the issuance of a "
        "Sanitary Certificate for our educational society located at 6-12-63, Namdevwada, Nizamabad-503002. "
        "We require the certificate for official purposes and request an inspection of our "
        "premises at your convenience."
    )
    story.append(Paragraph(body_text, body_style))

    body_text_2 = (
        "The institute maintains hygienic conditions, safe drinking water, and adequate sanitation "
        "facilities for students and staff. Kindly issue the Sanitary Certificate after inspection, "
        "similar to the standard certificate issued by the District Medical & Health Officer."
    )
    story.append(Paragraph(body_text_2, body_style))

    body_text_3 = (
        "We shall cooperate fully with the inspection team and provide any documents or information "
        "required."
    )
    story.append(Paragraph(body_text_3, body_style))

    story.append(Spacer(1, 12))

    story.append(Paragraph("Thank you.", body_style))
    story.append(Paragraph("Yours faithfully,", body_style))
    story.append(Spacer(1, 24))
    story.append(Paragraph("President", body_style))
    story.append(Paragraph("Prathibha Educational Society", body_style))

    doc.build(story)


if __name__ == "__main__":
    output_pdf = "Sanitary_Certificate_Request_Prathibha_Institute.pdf"
    build_request_letter(output_pdf)
    print(f"Generated: {output_pdf}")
