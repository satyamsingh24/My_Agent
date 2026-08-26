"""Render the CV template thumbnails used by the browser editor.

Run after changing CV_TEMPLATES so the editor's preview images stay in sync:
    python -m scripts.export_template_previews
"""
from __future__ import annotations

from pathlib import Path

from cv_engine import CV_TEMPLATES, build_template_preview, compose_cv_text

PREVIEW_DIR = Path(__file__).resolve().parent.parent / "static" / "editor" / "previews"
PREVIEW_WIDTH = 300

SAMPLE_SECTIONS = {
    "skills": "Python, Django, AWS, Docker, MySQL, Git, Linux",
    "experience": (
        "Example Technologies, Pune - Software Engineer (Jul 2023 - Present)\n"
        "Built and maintained backend REST APIs for a banking client\n"
        "Automated build and deployment steps with Jenkins pipelines"
    ),
    "projects": (
        "Expense Tracker - Django app with MySQL and Docker deployment\n"
        "Supports multi-user accounts and monthly reports"
    ),
    "education": "Master of Computer Application - RGPV Bhopal (2023 - 2025)",
    "certifications": "AWS Cloud Practitioner - Amazon (2024)",
}


def sample_text() -> str:
    return compose_cv_text(
        name="Your Name",
        email="you@example.com",
        phone="98765 43210",
        city="Indore",
        state="Madhya Pradesh",
        headline="Python Developer",
        sections=SAMPLE_SECTIONS,
    )


def main() -> None:
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    text = sample_text()
    for template in CV_TEMPLATES:
        image = build_template_preview(text, template, width=PREVIEW_WIDTH)
        path = PREVIEW_DIR / f"{template}.png"
        path.write_bytes(image)
        print(f"wrote {path.relative_to(PREVIEW_DIR.parent.parent.parent)}")


if __name__ == "__main__":
    main()
