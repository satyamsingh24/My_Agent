"""Convert a plain ATS resume .txt into a selectable, single-column PDF."""
from __future__ import annotations

import sys
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

HEADINGS = {
    "SUMMARY",
    "SKILLS",
    "PROFESSIONAL EXPERIENCE",
    "EDUCATION",
    "CERTIFICATIONS",
}


def txt_to_pdf(src: Path, dest: Path) -> None:
    styles = getSampleStyleSheet()
    name_style = ParagraphStyle(
        "Name",
        parent=styles["Normal"],
        fontName="Times-Bold",
        fontSize=16,
        leading=20,
        spaceAfter=2,
    )
    role_style = ParagraphStyle(
        "Role",
        parent=styles["Normal"],
        fontName="Times-Bold",
        fontSize=11,
        leading=14,
        spaceAfter=2,
    )
    contact_style = ParagraphStyle(
        "Contact",
        parent=styles["Normal"],
        fontName="Times-Roman",
        fontSize=10,
        leading=13,
        spaceAfter=8,
    )
    heading_style = ParagraphStyle(
        "Heading",
        parent=styles["Normal"],
        fontName="Times-Bold",
        fontSize=11,
        leading=14,
        spaceBefore=8,
        spaceAfter=4,
    )
    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontName="Times-Roman",
        fontSize=10,
        leading=13,
        spaceAfter=2,
    )
    bullet_style = ParagraphStyle(
        "Bullet",
        parent=body_style,
        leftIndent=12,
        bulletIndent=0,
        spaceAfter=2,
    )

    lines = src.read_text(encoding="utf-8").splitlines()
    story = []
    i = 0
    # Name, headline, contact
    if i < len(lines) and lines[i].strip():
        story.append(Paragraph(lines[i].strip(), name_style))
        i += 1
    if i < len(lines) and lines[i].strip():
        story.append(Paragraph(lines[i].strip(), role_style))
        i += 1
    if i < len(lines) and lines[i].strip():
        story.append(Paragraph(lines[i].strip().replace("|", " | "), contact_style))
        i += 1
    while i < len(lines) and not lines[i].strip():
        i += 1

    for raw in lines[i:]:
        line = raw.rstrip()
        if not line.strip():
            story.append(Spacer(1, 4))
            continue
        upper = line.strip().upper()
        if upper in HEADINGS:
            story.append(Paragraph(upper, heading_style))
            continue
        if line.strip().startswith(("- ", "• ")):
            text = line.strip()[2:].strip()
            story.append(Paragraph(f"• {text}", bullet_style))
            continue
        story.append(Paragraph(line.strip(), body_style))

    dest.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(dest),
        pagesize=letter,
        leftMargin=0.65 * inch,
        rightMargin=0.65 * inch,
        topMargin=0.55 * inch,
        bottomMargin=0.55 * inch,
        title=src.stem,
        author="Satyam Singh Bhadoriya",
    )
    doc.build(story)


if __name__ == "__main__":
    source = Path(sys.argv[1] if len(sys.argv) > 1 else "applications/Satyam_Singh_Bhadoriya_DevOps_ATS.txt")
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else source.with_suffix(".pdf")
    txt_to_pdf(source, output)
    print(output)
