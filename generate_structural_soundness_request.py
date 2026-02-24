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

    today = date(2026, 2, 6).strftime("%d-%m-%Y")

    story = []

    story.append(Paragraph("Request for Structural Soundness Certificate", title_style))
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
    story.append(Paragraph("The Executive Engineer (R&amp;B)", body_style))
    story.append(Paragraph("Division, Nizamabad", body_style))
    story.append(Paragraph("Government of Telangana", body_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Subject: Request for issue of Structural Soundness Certificate", body_style))
    story.append(Spacer(1, 4))

    story.append(Paragraph("Respected Sir/Madam,", body_style))

    body_text = (
        "I, the undersigned, on behalf of Prathibha Educational Society, request the issuance of a "
        "Structural Soundness Certificate for our educational society building located at 6-12-63, "
        "Namdevwada, Nizamabad-503002. We require the certificate for official purposes and request an "
        "inspection of the building at your convenience."
    )
    story.append(Paragraph(body_text, body_style))

    body_text_2 = (
        "The building is maintained in good condition and is being used for educational purposes. "
        "Kindly issue the Structural Soundness Certificate after inspection as per the standard format "
        "followed by the R&amp;B Department."
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
    output_pdf = "Structural_Soundness_Request_Prathibha_Educational_Society.pdf"
    build_request_letter(output_pdf)
    print(f"Generated: {output_pdf}")
