"""Core CV storage, PDF extraction, JD matching, and ATS PDF generation."""
from __future__ import annotations

import html
import json
import re
import secrets
from datetime import datetime
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from PIL import Image as PILImage

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "app_data"
CV_PDF = DATA_DIR / "existing_cv.pdf"
CV_TEXT = DATA_DIR / "existing_cv.txt"
CV_META = DATA_DIR / "existing_cv.json"
OUTPUT_DIR = ROOT / "applications" / "generated"

SKILLS = [
    "AWS", "GCP", "Azure", "GitHub Actions", "Infrastructure as Code",
    "CI/CD", "CloudFormation",
    "Kubernetes", "Terraform", "Jenkins", "Docker", "Ansible", "Helm",
    "Prometheus", "Grafana", "Datadog", "CloudWatch", "OpenTelemetry",
    "Python", "Java", "JavaScript", "TypeScript", "C#", "C++", "Go",
    "Ruby", "PHP", "Kotlin", "Swift", "Scala", "Dart", "R", "Shell", "Bash",
    "PowerShell", "HTML", "CSS", "Sass", "Bootstrap", "Tailwind CSS",
    "React", "Angular", "Vue", "Next.js", "Node.js", "Express.js", "Redux",
    "Django", "Flask", "FastAPI", "Spring Boot", "Spring", "Hibernate",
    "JPA", "JDBC", "Servlets", "JSP", ".NET", "ASP.NET", "REST API",
    "GraphQL", "WebSocket", "JWT", "OAuth", "Microservices",
    "Flutter", "React Native", "Android", "iOS",
    "Linux", "Ubuntu", "Windows", "Git", "GitHub", "GitLab",
    "Bitbucket", "MySQL", "PostgreSQL", "MongoDB", "Redis", "Oracle",
    "DynamoDB", "SQLite", "SQL Server", "SQL", "NoSQL", "Kafka", "RabbitMQ",
    "Nginx", "Apache",
    "Tomcat", "EC2", "IAM", "S3", "RDS", "Lambda", "Route53", "SQS", "SNS",
    "KMS", "VPC", "VPN", "CDN",
    "Agile", "Scrum", "Jira", "Selenium", "Pytest", "JUnit", "SonarQube",
    "Cypress", "Playwright", "Postman", "Maven", "Gradle", "npm", "Yarn",
    "Machine Learning", "Artificial Intelligence", "Generative AI", "LLM",
    "RAG", "Data Science", "Data Engineering", "Pandas", "NumPy", "SciPy",
    "Scikit-learn", "TensorFlow", "PyTorch", "Keras", "Spark", "Hadoop",
    "Airflow", "Databricks", "Power BI", "Tableau", "Excel",
    "Cybersecurity", "DevSecOps", "OWASP", "UI/UX", "Figma",
]

SECTION_HEADINGS = {
    "SUMMARY", "PROFILE", "PROFESSIONAL SUMMARY", "OBJECTIVE", "SKILLS",
    "TECHNICAL SKILLS", "CORE SKILLS", "EXPERIENCE", "WORK EXPERIENCE",
    "PROFESSIONAL EXPERIENCE", "PROJECTS", "EDUCATION", "CERTIFICATIONS",
    "CERTIFICATES", "ACHIEVEMENTS", "LANGUAGES", "INTERNSHIP", "INTERNSHIPS",
    "CORE COMPETENCIES", "KEY COMPETENCIES", "AREAS OF EXPERTISE",
}

# Canonical structure copied from profile/Satyam_Dev_Resume_ATS.pdf so every
# generated CV follows the same reference layout.
HEADING_ALIASES = {
    "SUMMARY": "SUMMARY",
    "PROFILE": "SUMMARY",
    "OBJECTIVE": "SUMMARY",
    "PROFESSIONAL SUMMARY": "SUMMARY",
    "SKILLS": "SKILLS",
    "TECHNICAL SKILLS": "SKILLS",
    "CORE SKILLS": "SKILLS",
    "CORE COMPETENCIES": "CORE COMPETENCIES",
    "KEY COMPETENCIES": "CORE COMPETENCIES",
    "AREAS OF EXPERTISE": "CORE COMPETENCIES",
    "JD-MATCHED SKILLS": "JD-MATCHED SKILLS",
    "EXPERIENCE": "PROFESSIONAL EXPERIENCE",
    "WORK EXPERIENCE": "PROFESSIONAL EXPERIENCE",
    "PROFESSIONAL EXPERIENCE": "PROFESSIONAL EXPERIENCE",
    "INTERNSHIP": "INTERNSHIPS",
    "INTERNSHIPS": "INTERNSHIPS",
    "PROJECTS": "PROJECTS",
    "EDUCATION": "EDUCATION",
    "CERTIFICATIONS": "CERTIFICATIONS",
    "CERTIFICATES": "CERTIFICATIONS",
    "ACHIEVEMENTS": "ACHIEVEMENTS",
    "LANGUAGES": "LANGUAGES",
    "ADDITIONAL DETAILS": "ADDITIONAL DETAILS",
}

REFERENCE_ORDER = [
    "SUMMARY",
    "SKILLS",
    "JD-MATCHED SKILLS",
    "CORE COMPETENCIES",
    "PROFESSIONAL EXPERIENCE",
    "INTERNSHIPS",
    "PROJECTS",
    "EDUCATION",
    "CERTIFICATIONS",
    "ACHIEVEMENTS",
    "LANGUAGES",
    "ADDITIONAL DETAILS",
]

# ATS parsers gain nothing from these blocks in the source CV.
DROP_HEADINGS = {"DECLARATION", "REFERENCES", "DECLARATION:"}

# Sections whose plain lines are prose, not company/degree entry lines.
PROSE_SECTIONS = {"SUMMARY", "SKILLS", "JD-MATCHED SKILLS", "LANGUAGES"}

# Skill buckets used to print a grouped, recruiter-friendly SKILLS block.
SKILL_GROUPS: list[tuple[str, tuple[str, ...]]] = [
    ("Cloud Platforms", (
        "aws", "amazon web services", "azure", "gcp", "google cloud", "ec2",
        "s3", "iam", "rds", "lambda", "route53", "route 53", "sqs", "sns",
        "kms", "vpc", "cloudformation", "cloud deployment", "cloud",
    )),
    ("DevOps & Automation", (
        "ci/cd", "cicd", "infrastructure as code", "iac", "terraform",
        "jenkins", "docker", "kubernetes", "ansible", "helm",
        "github actions", "gitlab ci", "configuration management",
        "infrastructure automation", "automation", "deployment", "devsecops",
    )),
    ("Monitoring & Observability", (
        "prometheus", "grafana", "datadog", "cloudwatch", "opentelemetry",
        "monitoring", "logging",
    )),
    ("Programming & Scripting", (
        "java", "python", "javascript", "typescript", "c#", "c++", "go",
        "golang", "kotlin", "swift", "scala", "ruby", "php", "dart", "shell",
        "bash", "powershell",
    )),
    ("Frameworks & Libraries", (
        "spring boot", "spring", "hibernate", "jpa", "jdbc", "servlets",
        "jsp", "django", "flask", "fastapi", "react", "angular", "vue",
        "next.js", "node.js", "express.js", "redux", ".net", "asp.net",
        "rest api", "graphql", "microservices", "html", "css", "bootstrap",
        "tailwind", "flutter", "react native",
    )),
    ("Databases", (
        "mysql", "postgresql", "mongodb", "redis", "oracle", "dynamodb",
        "sqlite", "sql", "nosql", "kafka", "rabbitmq",
    )),
    ("Tools & Platforms", (
        "git", "github", "gitlab", "bitbucket", "linux", "ubuntu", "windows",
        "maven", "gradle", "npm", "yarn", "jira", "postman", "nginx",
        "apache", "tomcat", "figma", "development tools",
    )),
    ("Data & AI", (
        "machine learning", "artificial intelligence", "generative ai", "llm",
        "rag", "data science", "data engineering", "pandas", "numpy",
        "scikit-learn", "tensorflow", "pytorch", "spark", "hadoop", "airflow",
        "power bi", "tableau", "excel",
    )),
    ("Practices", (
        "agile", "scrum", "owasp", "cybersecurity", "ui/ux", "testing",
        "unit testing", "code review",
    )),
]

# Professional phrasing for competency bullets built from the user's own words.
COMPETENCY_TEMPLATES: list[tuple[tuple[str, ...], str]] = [
    (("infrastructure as code", "iac", "terraform", "cloudformation"),
     "Provisioning and managing infrastructure using Infrastructure as Code "
     "(IaC) for repeatable, version-controlled environments."),
    (("ci/cd", "cicd", "jenkins", "github actions", "pipeline"),
     "Building and maintaining CI/CD pipelines that automate build, test, and "
     "deployment stages."),
    (("configuration management", "ansible", "puppet", "chef"),
     "Standardising server and application configuration to keep environments "
     "consistent and reproducible."),
    (("infrastructure automation", "automation", "scripting", "shell", "bash"),
     "Automating repetitive build, deployment, and operational tasks to reduce "
     "manual effort and human error."),
    (("cloud deployment", "deployment process", "deployment", "release"),
     "Planning and executing application deployments with repeatable, "
     "low-downtime release processes."),
    (("aws", "azure", "gcp", "cloud"),
     "Working with cloud platform services for compute, storage, networking, "
     "and access management."),
    (("docker", "container"),
     "Containerising applications for portable builds across development and "
     "production environments."),
    (("kubernetes", "helm", "orchestration"),
     "Deploying and operating containerised workloads on Kubernetes, including "
     "configuration and scaling."),
    (("monitoring", "prometheus", "grafana", "datadog", "cloudwatch"),
     "Setting up monitoring and alerting to track system health and respond to "
     "issues early."),
    (("git", "github", "gitlab", "version control"),
     "Managing source code with Git-based version control, branching, and pull "
     "request reviews."),
    (("linux", "ubuntu", "server administration"),
     "Administering Linux servers, including packages, permissions, services, "
     "and troubleshooting."),
    (("java", "spring", "hibernate", "j2ee"),
     "Developing and maintaining Java applications using object-oriented "
     "design and framework-based development."),
    (("python", "django", "flask", "fastapi"),
     "Developing Python services and automation scripts with clean, "
     "maintainable code."),
    (("react", "angular", "vue", "javascript", "typescript", "frontend"),
     "Building responsive front-end interfaces and integrating them with "
     "backend APIs."),
    (("rest api", "api", "microservices", "backend"),
     "Designing and consuming REST APIs and service integrations."),
    (("sql", "mysql", "postgresql", "mongodb", "database"),
     "Designing database schemas and writing optimised queries for "
     "application data access."),
    (("development tools", "tooling", "infrastructure setup"),
     "Setting up development tooling and project infrastructure to support "
     "smooth team delivery."),
    (("agile", "scrum", "sprint"),
     "Working in Agile/Scrum teams with sprint planning, reviews, and "
     "collaborative delivery."),
    (("testing", "junit", "pytest", "selenium", "qa"),
     "Writing and running automated tests to protect functionality during "
     "changes."),
]

MAX_COMPETENCY_BULLETS = 8

# Words that already read as a job title, so the summary needs no suffix.
ROLE_NOUNS = (
    "developer", "engineer", "analyst", "administrator", "architect",
    "consultant", "manager", "specialist", "designer", "tester", "intern",
    "trainee", "lead", "scientist", "programmer", "sre", "devops", "support",
)


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract selectable text from an uploaded PDF."""
    reader = PdfReader(BytesIO(pdf_bytes))
    pages = [(page.extract_text() or "").strip() for page in reader.pages]
    text = "\n".join(page for page in pages if page)
    text = clean_text(text)
    if len(text) < 80:
        raise ValueError(
            "No readable text was found in this PDF. Please upload a "
            "selectable-text CV instead of a scanned image PDF."
        )
    return text


BULLET_GLYPHS = "\u2022\u2023\u2043\u2219\u25aa\u25cf\u25e6\u00b7\u007f\uf0a7\uf0b7\uf06c"


def clean_text(text: str) -> str:
    for glyph in BULLET_GLYPHS:
        text = text.replace(glyph, "-")
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = text.replace("\xa0", " ")
    lines: list[str] = []
    blank = False
    for raw in text.splitlines():
        line = re.sub(r"[ \t]+", " ", raw).strip()
        if not line:
            if lines and not blank:
                lines.append("")
            blank = True
            continue
        blank = False
        lines.append(line)
    return "\n".join(lines).strip()


def save_existing_cv(pdf_bytes: bytes, original_name: str, text: str) -> dict:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CV_PDF.write_bytes(pdf_bytes)
    CV_TEXT.write_text(text, encoding="utf-8")
    meta = {
        "original_name": original_name,
        "uploaded_at": datetime.now().isoformat(timespec="seconds"),
        "characters": len(text),
    }
    CV_META.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return meta


def load_existing_cv() -> tuple[dict, str, bytes] | None:
    if not (CV_PDF.exists() and CV_TEXT.exists() and CV_META.exists()):
        return None
    try:
        meta = json.loads(CV_META.read_text(encoding="utf-8"))
        return meta, CV_TEXT.read_text(encoding="utf-8"), CV_PDF.read_bytes()
    except (OSError, ValueError, json.JSONDecodeError):
        return None


def _contains(text: str, term: str) -> bool:
    aliases = {
        "AWS": ["amazon web services", "aws"],
        "Azure": ["microsoft azure", "azure"],
        "GCP": ["google cloud platform", "gcp"],
        "Go": ["golang", "go"],
        "Route53": ["route 53", "route53"],
        "Artificial Intelligence": ["artificial intelligence", "ai"],
        "Machine Learning": ["machine learning", "ml"],
        "Generative AI": ["generative ai", "genai", "gen ai"],
        "LLM": ["llm", "large language model", "large language models"],
        "Kubernetes": ["kubernetes", "k8s"],
        "Infrastructure as Code": ["infrastructure as code", "iac"],
        "Node.js": ["node.js", "nodejs", "node js"],
        "Next.js": ["next.js", "nextjs", "next js"],
        "Express.js": ["express.js", "expressjs", "express js"],
        "JavaScript": ["javascript", "java script", "js"],
        "TypeScript": ["typescript", "type script", "ts"],
        "REST API": ["rest api", "restful api", "restful services"],
        "Spring Boot": ["spring boot", "springboot"],
        "PostgreSQL": ["postgresql", "postgres"],
        "Scikit-learn": ["scikit-learn", "sklearn"],
        "Power BI": ["power bi", "powerbi"],
        "HTML": ["html", "html5"],
        "CSS": ["css", "css3"],
    }
    options = aliases.get(term, [term])
    return any(
        re.search(r"(?<!\w)" + re.escape(option) + r"(?!\w)", text, re.IGNORECASE)
        for option in options
    )


def match_jd(cv_text: str, jd_text: str) -> dict:
    requested = [skill for skill in SKILLS if _contains(jd_text, skill)]
    present = [skill for skill in requested if _contains(cv_text, skill)]
    missing = [skill for skill in requested if skill not in present]
    score = round(100 * len(present) / len(requested)) if requested else 0
    return {
        "requested": _dedupe(requested),
        "matched": _dedupe(present),
        "missing": _dedupe(missing),
        "score": score,
    }


def _dedupe(items: list[str]) -> list[str]:
    seen: set[str] = set()
    result = []
    for item in items:
        key = item.casefold()
        if key not in seen:
            seen.add(key)
            result.append(item)
    return result


def build_tailored_text(cv_text: str, match: dict) -> str:
    """Preserve source CV, add truthful JD keywords, and apply reference layout."""
    matched = match["matched"]
    keyword_line = (
        ", ".join(matched)
        if matched
        else "No verified JD keywords found in the uploaded CV."
    )
    combined = cv_text.rstrip() + "\n\nJD-MATCHED SKILLS\n" + keyword_line + "\n"
    return restructure_to_reference(combined)


def parse_cv_sections(text: str) -> tuple[list[str], dict[str, list[str]], list[str]]:
    """Split any CV text into header lines and canonical sections."""
    lines = clean_text(text).splitlines()
    header: list[str] = []
    sections: dict[str, list[str]] = {}
    order: list[str] = []
    current: str | None = None
    dropping = False

    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        key = line.upper().rstrip(":")
        if key in DROP_HEADINGS:
            dropping = True
            current = None
            continue
        canonical = HEADING_ALIASES.get(key)
        if canonical:
            dropping = False
            current = canonical
            if canonical not in sections:
                sections[canonical] = []
                order.append(canonical)
            continue
        if dropping:
            continue
        if current is None:
            header.append(line)
        else:
            sections[current].append(line)

    return header, sections, order


def restructure_to_reference(text: str) -> str:
    """Reorder any CV text into the reference resume layout, ATS-safe."""
    header, sections, order = parse_cv_sections(text)

    out: list[str] = header[:3] if header else []
    leftover_header = header[3:]
    if leftover_header:
        sections.setdefault("ADDITIONAL DETAILS", [])
        if "ADDITIONAL DETAILS" not in order:
            order.append("ADDITIONAL DETAILS")
        sections["ADDITIONAL DETAILS"] = leftover_header + sections["ADDITIONAL DETAILS"]

    ordered = [name for name in REFERENCE_ORDER if sections.get(name)]
    ordered += [name for name in order if name not in ordered and sections.get(name)]

    for name in ordered:
        out += ["", name]
        out += sections[name]

    return "\n".join(out).strip() + "\n"


def detect_skills(text: str) -> list[str]:
    """Return known skills that actually appear in the given text."""
    return _dedupe([skill for skill in SKILLS if _contains(text, skill)])


ROLE_WORDS = (
    "developer", "engineer", "analyst", "administrator", "architect",
    "consultant", "manager", "specialist", "designer", "tester", "trainee",
    "intern", "lead", "scientist", "programmer", "devops", "support",
)


def _cv_identity(text: str) -> tuple[str, str, str, str]:
    """Extract only identity fields that are visibly present in the CV."""
    header, sections, _ = parse_cv_sections(text)
    all_lines = [line.strip() for line in clean_text(text).splitlines() if line.strip()]
    candidates = header or all_lines[:8]

    email_match = re.search(
        r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", text, re.IGNORECASE
    )
    phone_match = re.search(
        r"(?<!\d)(?:\+?\d[\d ()-]{7,}\d)(?!\d)", text
    )
    email = email_match.group(0) if email_match else ""
    phone = re.sub(r"\s+", " ", phone_match.group(0)).strip() if phone_match else ""

    def is_contact(line: str) -> bool:
        low = line.lower()
        return (
            "@" in line
            or bool(re.search(r"\d{7,}", re.sub(r"\D", "", line)))
            or "linkedin" in low
            or "github" in low
            or "|" in line
        )

    name = next(
        (
            line
            for line in candidates
            if not is_contact(line)
            and line.upper().rstrip(":") not in SECTION_HEADINGS
            and line.upper() not in {"RESUME", "CURRICULUM VITAE", "CV"}
            and 1 <= len(line.split()) <= 6
        ),
        "Candidate",
    )

    role = next(
        (
            line
            for line in candidates
            if line != name
            and not is_contact(line)
            and any(word in line.lower() for word in ROLE_WORDS)
            and len(line.split()) <= 10
        ),
        "",
    )
    if not role:
        summary = " ".join(sections.get("SUMMARY", [])[:3])
        title_match = re.search(
            r"\b([A-Za-z+#.]+(?:\s+[A-Za-z+#.]+){0,4}\s+"
            r"(?:Developer|Engineer|Analyst|Administrator|Architect|"
            r"Consultant|Specialist|Designer|Tester|Manager))\b",
            summary,
            re.IGNORECASE,
        )
        role = title_match.group(1) if title_match else ""

    return name.title() if name.isupper() else name, role, email, phone


def build_cold_email(
    cv_text: str,
    target_role: str = "",
    company: str = "",
) -> dict[str, str | list[str]]:
    """Create a concise, truthful HR email using only facts found in the CV."""
    header, sections, _ = parse_cv_sections(cv_text)
    del header
    name, inferred_role, email, phone = _cv_identity(cv_text)
    role = target_role.strip() or inferred_role.strip()
    if role and (role.isupper() or role.islower()):
        role = role.title()
    role = re.sub(r"\bDevops\b", "DevOps", role)
    role_label = role or "relevant"
    company_name = company.strip()

    experience_lines = [
        line.strip()
        for line in sections.get("PROFESSIONAL EXPERIENCE", [])
        if line.strip()
    ]
    internship_lines = [
        line.strip()
        for line in sections.get("INTERNSHIPS", [])
        if line.strip()
    ]
    has_experience = bool(experience_lines)
    has_internship = bool(internship_lines)

    years_match = re.search(
        r"\b(\d+(?:\.\d+)?\+?)\s*(?:years?|yrs?)\b",
        " ".join(sections.get("SUMMARY", [])),
        re.IGNORECASE,
    )
    years = years_match.group(1) if years_match else ""
    skills = detect_skills(cv_text)[:6]
    skill_phrase = ", ".join(skills)

    if has_experience:
        status = "Experienced professional"
        experience_phrase = (
            f"I have {years} years of professional experience"
            if years
            else "I have professional experience"
        )
        if role:
            experience_phrase += f" as a {role}"
        experience_phrase += "."
    elif has_internship:
        status = "Fresher with internship experience"
        experience_phrase = (
            f"I am a fresher pursuing {role_label} opportunities, with practical "
            "internship exposure documented in my CV."
        )
    else:
        status = "Fresher"
        experience_phrase = (
            f"I am a fresher seeking {role_label} opportunities and looking to "
            "apply the skills and project knowledge documented in my CV."
        )

    if skills:
        background = f"My core skills include {skill_phrase}."
    else:
        background = (
            "My attached CV outlines my relevant education, projects, and skills."
        )

    destination = f" at {company_name}" if company_name else " in your organization"
    subject_role = role if role else "Relevant"
    subject = f"Application for {subject_role} Opportunities"
    if name != "Candidate":
        subject += f" | {name}"

    signoff = [name]
    if phone:
        signoff.append(phone)
    if email and email.casefold() != name.casefold():
        signoff.append(email)

    body = "\n".join(
        [
            "Dear Hiring Manager,",
            "",
            "I hope you are doing well.",
            "",
            f"My name is {name}. {experience_phrase} {background}",
            "",
            f"I am reaching out to explore suitable {role_label} opportunities"
            f"{destination}. I have attached my CV for your review and would "
            "appreciate the opportunity to discuss how my background could "
            "contribute to your team.",
            "",
            "Thank you for your time and consideration.",
            "",
            "Best regards,",
            *signoff,
        ]
    )

    return {
        "subject": subject,
        "body": body,
        "status": status,
        "name": name,
        "role": role,
        "skills": skills,
    }


def _section_lines(heading: str, content: str) -> list[str]:
    """Keep user's own bullets as bullets; plain lines stay entry/detail lines."""
    text = clean_text(content)
    if not text:
        return []
    rows = ["", heading]
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith(("-", "*", "•")):
            rows.append("- " + line.lstrip("-*• ").strip())
        elif line.upper().rstrip(":") in SECTION_HEADINGS:
            rows += ["", line.upper().rstrip(":")]
        else:
            rows.append(line)
    return rows


def _looks_like_skill_list(lines: list[str]) -> bool:
    """True when the free-text lines read like a skill list, not sentences."""
    body = [line for line in lines if line]
    if not body:
        return False
    if len(body) == 1:
        only = body[0].rstrip()
        return not only.endswith(".") and ("," in only or len(only.split()) <= 6)
    short = [
        line
        for line in body
        if len(line.split()) <= 6 and not line.rstrip().endswith(".")
    ]
    return len(short) >= max(2, int(len(body) * 0.7))


def group_skills(items: list[str]) -> list[str]:
    """Print skills as labelled groups so the block reads like a pro resume."""
    if len(items) < 4:
        return [", ".join(items)] if items else []
    buckets: dict[str, list[str]] = {}
    leftover: list[str] = []
    for item in items:
        low = item.lower()
        for label, keywords in SKILL_GROUPS:
            if any(word in low for word in keywords):
                buckets.setdefault(label, []).append(item)
                break
        else:
            leftover.append(item)
    rows = [
        f"{label}: {', '.join(buckets[label])}"
        for label, _ in SKILL_GROUPS
        if buckets.get(label)
    ]
    if leftover:
        rows.append(f"Additional Skills: {', '.join(leftover)}")
    return rows


def build_competencies(phrases: list[str], skills: list[str]) -> list[str]:
    """Turn the user's own skill phrases into professional competency bullets."""
    source = [phrase.strip() for phrase in phrases + skills if phrase.strip()]
    if not source:
        return []
    bullets: list[str] = []
    for keywords, sentence in COMPETENCY_TEMPLATES:
        if sentence in bullets:
            continue
        if any(word in phrase.lower() for phrase in source for word in keywords):
            bullets.append(sentence)
        if len(bullets) >= MAX_COMPETENCY_BULLETS:
            break
    return bullets


def build_summary(headline: str, skills: list[str], has_history: bool) -> list[str]:
    """Write a professional summary when the user leaves that box empty."""
    role = headline.strip()
    if role and role == role.lower():
        role = role.title()
    if role and not any(noun in role.lower() for noun in ROLE_NOUNS):
        role = f"{role} professional"
    role = role or "Technology professional"
    top = [skill for skill in skills if len(skill) < 40][:5]
    sentences = []
    if has_history:
        sentences.append(
            f"{role} with hands-on project and workplace experience delivering "
            "reliable, well-documented solutions."
        )
    else:
        sentences.append(
            f"{role} with practical, hands-on exposure to modern development "
            "and deployment workflows."
        )
    if top:
        sentences.append(f"Core strengths across {', '.join(top)}.")
    sentences.append(
        "Comfortable owning work end to end — from requirement and design "
        "through implementation, testing, and deployment — and improving "
        "delivery through automation and clear documentation."
    )
    sentences.append(
        "Quick to pick up new tools and keen to contribute in a collaborative, "
        "delivery-focused team."
    )
    return sentences


def compose_cv_text(
    name: str = "",
    email: str = "",
    phone: str = "",
    address: str = "",
    city: str = "",
    state: str = "",
    details: str = "",
    sections: dict[str, str] | None = None,
    headline: str = "",
) -> str:
    """Build an ATS-safe CV from whatever the user filled. No field is required."""
    location = ", ".join(part.strip() for part in (city, state) if part.strip())
    contact_parts = [phone, email, address, location]
    contact = " | ".join(part.strip() for part in contact_parts if part.strip())

    lines = [(name.strip() or "CURRICULUM VITAE").upper()]
    if headline.strip():
        lines.append(headline.strip().upper())
    if contact:
        lines.append(contact)

    filled = sections or {}
    extra = details.strip()

    extra_headed: list[str] = []
    extra_loose: list[str] = []
    current_extra = ""
    for raw in clean_text(extra).splitlines():
        line = raw.strip()
        if not line:
            continue
        upper = line.upper().rstrip(":")
        if upper in SECTION_HEADINGS:
            current_extra = upper
            extra_headed += ["", upper]
            continue
        if line.startswith(("-", "*", "•")):
            line = "- " + line.lstrip("-*• ").strip()
        if current_extra:
            extra_headed.append(line)
        else:
            extra_loose.append(line)

    loose_as_skills = _looks_like_skill_list(
        [line for line in extra_loose if not line.startswith("- ")]
    )

    skill_items: list[str] = []
    if filled.get("skills"):
        skill_items += [
            part.strip()
            for part in re.split(r"[\n,;]+", filled["skills"])
            if part.strip()
        ]
    if loose_as_skills:
        for line in extra_loose:
            skill_items += [
                part.strip()
                for part in re.split(r"[,;]|\s+/\s+", line.lstrip("- "))
                if part.strip()
            ]
    known = detect_skills("\n".join([extra, *filled.values()]))
    joined = " ".join(skill_items).lower()
    skill_items += [skill for skill in known if skill.lower() not in joined]
    skill_items = _dedupe(skill_items)

    has_history = any(
        filled.get(key, "").strip()
        for key in ("experience", "internships", "projects")
    )
    if filled.get("summary"):
        lines += ["", "SUMMARY"]
        lines += [
            line.strip()
            for line in clean_text(filled["summary"]).splitlines()
            if line.strip()
        ]
    elif skill_items or headline.strip():
        lines += ["", "SUMMARY"]
        lines += build_summary(headline, skill_items, has_history)

    if skill_items:
        lines += ["", "SKILLS"]
        lines += group_skills(skill_items)

    competency_source = extra_loose if loose_as_skills else []
    competencies = build_competencies(competency_source, skill_items)
    if competencies and not has_history:
        lines += ["", "CORE COMPETENCIES"]
        lines += ["- " + bullet for bullet in competencies]

    for heading, key in (
        ("PROFESSIONAL EXPERIENCE", "experience"),
        ("PROJECTS", "projects"),
        ("EDUCATION", "education"),
        ("CERTIFICATIONS", "certifications"),
        ("INTERNSHIPS", "internships"),
        ("ACHIEVEMENTS", "achievements"),
        ("LANGUAGES", "languages"),
    ):
        lines += _section_lines(heading, filled.get(key, ""))

    lines += extra_headed
    if extra_loose and not loose_as_skills:
        lines += ["", "ADDITIONAL DETAILS"]
        lines += [
            line if line.startswith("- ") else "- " + line for line in extra_loose
        ]

    return restructure_to_reference("\n".join(lines))


def prepare_profile_photo(photo_bytes: bytes) -> BytesIO:
    """Crop a portrait photo to a consistent size for the CV header."""
    try:
        image = PILImage.open(BytesIO(photo_bytes))
        image = image.convert("RGB")
    except Exception as exc:
        raise ValueError(
            "Could not read the profile photo. Please upload a JPG, PNG, or WEBP image."
        ) from exc

    target_w, target_h = 340, 420
    src_w, src_h = image.size
    if src_w < 40 or src_h < 40:
        raise ValueError("Profile photo is too small. Please upload a clearer picture.")

    scale = max(target_w / src_w, target_h / src_h)
    resized = image.resize(
        (max(1, int(src_w * scale)), max(1, int(src_h * scale))),
        PILImage.Resampling.LANCZOS,
    )
    left = max(0, (resized.width - target_w) // 2)
    top = max(0, (resized.height - target_h) // 2)
    cropped = resized.crop((left, top, left + target_w, top + target_h))
    output = BytesIO()
    cropped.save(output, format="JPEG", quality=88, optimize=True)
    output.seek(0)
    return output


def _split_header_lines(text: str) -> tuple[list[str], list[str]]:
    """Take the name / role / contact block off the top of the CV."""
    header: list[str] = []
    rest: list[str] = []
    in_body = False
    for raw in text.splitlines():
        line = raw.strip()
        if not in_body:
            if not line:
                continue
            upper = line.upper().rstrip(":")
            if upper in SECTION_HEADINGS or upper == "JD-MATCHED SKILLS" or len(header) >= 3:
                in_body = True
                rest.append(raw)
                continue
            header.append(line)
            continue
        rest.append(raw)
    return header, rest


def random_output_stem() -> str:
    """File name for downloads: My_Agent + 5 random digits."""
    return f"My_Agent_{secrets.randbelow(90000) + 10000}"


def manual_output_stem(name: str) -> str:
    base = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_") or "My"
    return f"{base}_ATS_CV"


def build_ats_pdf(
    text: str,
    title: str = "ATS Resume",
    photo_bytes: bytes | None = None,
) -> bytes:
    """Create an attractive, selectable, one-column ATS-safe PDF."""
    output = BytesIO()
    styles = getSampleStyleSheet()
    navy = colors.HexColor("#28306B")
    pale_blue = colors.HexColor("#EEF0FA")
    charcoal = colors.HexColor("#2A2E3A")
    name_style = ParagraphStyle(
        "Name", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=17, leading=20, spaceAfter=4, textColor=navy,
    )
    heading_style = ParagraphStyle(
        "Heading", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=10.5, leading=13, spaceBefore=8, spaceAfter=4,
        textColor=navy, backColor=pale_blue, borderPadding=(3, 5, 3, 5),
    )
    body_style = ParagraphStyle(
        "Body", parent=styles["Normal"], fontName="Helvetica",
        fontSize=9.5, leading=12.2, spaceAfter=2, textColor=charcoal,
    )
    bullet_style = ParagraphStyle(
        "Bullet", parent=body_style, leftIndent=13, firstLineIndent=-9,
        bulletIndent=1, spaceAfter=2,
    )
    role_style = ParagraphStyle(
        "Role", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=11, leading=13.5, spaceAfter=3,
        textColor=colors.HexColor("#5A4B9C"),
    )
    entry_style = ParagraphStyle(
        "Entry", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=10, leading=12.5, spaceBefore=3, spaceAfter=1,
        textColor=colors.HexColor("#1F2437"),
    )

    story = []
    source_lines = text.splitlines()
    nonempty_seen = 0
    if photo_bytes:
        header, rest = _split_header_lines(text)
        left = []
        for index, line in enumerate(header):
            safe = html.escape(line)
            if index == 0:
                left.append(Paragraph(safe, name_style))
            elif index == 1 and line == line.upper() and len(line.split()) <= 6:
                left.append(Paragraph(safe, role_style))
            else:
                left.append(Paragraph(safe, body_style))
        if not left:
            left = [Paragraph(" ", body_style)]
        photo = RLImage(
            prepare_profile_photo(photo_bytes),
            width=1.12 * inch,
            height=1.38 * inch,
        )
        table = Table([[left, photo]], colWidths=[5.78 * inch, 1.22 * inch])
        table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(table)
        story.append(Spacer(1, 6))
        source_lines = rest
        nonempty_seen = 3

    section = ""
    for raw in source_lines:
        line = raw.strip()
        if not line:
            story.append(Spacer(1, 3))
            continue
        safe = html.escape(line)
        upper = line.upper().rstrip(":")
        if nonempty_seen == 0:
            story.append(Paragraph(safe, name_style))
        elif nonempty_seen == 1 and line == upper and len(line.split()) <= 6:
            story.append(Paragraph(safe, role_style))
        elif upper in SECTION_HEADINGS or upper == "JD-MATCHED SKILLS":
            section = HEADING_ALIASES.get(upper, upper)
            story.append(Paragraph(html.escape(upper), heading_style))
        elif line.startswith(("-", "*", "•")):
            content = html.escape(line.lstrip("-*• ").strip())
            story.append(Paragraph(f"• {content}", bullet_style))
        elif section and section not in PROSE_SECTIONS:
            story.append(Paragraph(safe, entry_style))
        else:
            story.append(Paragraph(safe, body_style))
        nonempty_seen += 1

    doc = SimpleDocTemplate(
        output, pagesize=A4, leftMargin=0.65 * inch, rightMargin=0.65 * inch,
        topMargin=0.55 * inch, bottomMargin=0.55 * inch,
        title=title, author="My_AGENT - Satyam Singh Bhadoriya",
    )
    doc.build(story)
    return output.getvalue()


def _set_paragraph_bottom_border(paragraph) -> None:
    """Add a subtle ATS-safe border below a DOCX section heading."""
    p_pr = paragraph._p.get_or_add_pPr()
    borders = p_pr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        p_pr.append(borders)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "6A6FB5")
    borders.append(bottom)


def _clear_table_borders(table) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        item = OxmlElement(f"w:{edge}")
        item.set(qn("w:val"), "nil")
        item.set(qn("w:sz"), "0")
        item.set(qn("w:space"), "0")
        item.set(qn("w:color"), "auto")
        borders.append(item)
    tbl_pr.append(borders)


def build_ats_docx(
    text: str,
    title: str = "ATS Resume",
    photo_bytes: bytes | None = None,
) -> bytes:
    """Create a single-column DOCX preferred by many ATS platforms."""
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.05

    source_lines = text.splitlines()
    nonempty_seen = 0
    if photo_bytes:
        header, rest = _split_header_lines(text)
        table = document.add_table(rows=1, cols=2)
        _clear_table_borders(table)
        left_cell, right_cell = table.cell(0, 0), table.cell(0, 1)
        left_cell.width = Inches(5.8)
        right_cell.width = Inches(1.25)
        for index, line in enumerate(header):
            paragraph = (
                left_cell.paragraphs[0] if index == 0 else left_cell.add_paragraph()
            )
            run = paragraph.add_run(line)
            run.font.name = "Arial"
            if index == 0:
                run.bold = True
                run.font.size = Pt(18)
                run.font.color.rgb = RGBColor(0x28, 0x30, 0x6B)
            elif index == 1 and line == line.upper() and len(line.split()) <= 6:
                run.bold = True
                run.font.size = Pt(11.5)
                run.font.color.rgb = RGBColor(0x5A, 0x4B, 0x9C)
            else:
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0x2A, 0x2E, 0x3A)
            paragraph.paragraph_format.space_after = Pt(3)
        photo_para = right_cell.paragraphs[0]
        photo_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        photo_para.add_run().add_picture(
            prepare_profile_photo(photo_bytes), width=Inches(1.12)
        )
        source_lines = rest
        nonempty_seen = 3

    section = ""
    for raw in source_lines:
        line = raw.strip()
        if not line:
            document.add_paragraph().paragraph_format.space_after = Pt(1)
            continue
        upper = line.upper().rstrip(":")
        if nonempty_seen == 0:
            paragraph = document.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = paragraph.add_run(line)
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(0x28, 0x30, 0x6B)
            paragraph.paragraph_format.space_after = Pt(3)
        elif nonempty_seen == 1 and line == upper and len(line.split()) <= 6:
            paragraph = document.add_paragraph()
            run = paragraph.add_run(line)
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(0x5A, 0x4B, 0x9C)
            paragraph.paragraph_format.space_after = Pt(3)
        elif upper in SECTION_HEADINGS or upper == "JD-MATCHED SKILLS":
            section = HEADING_ALIASES.get(upper, upper)
            paragraph = document.add_paragraph()
            run = paragraph.add_run(upper)
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0x28, 0x30, 0x6B)
            paragraph.paragraph_format.space_before = Pt(7)
            paragraph.paragraph_format.space_after = Pt(3)
            _set_paragraph_bottom_border(paragraph)
        elif line.startswith(("-", "*", "•")):
            paragraph = document.add_paragraph(style="List Bullet")
            paragraph.add_run(line.lstrip("-*• ").strip())
            paragraph.paragraph_format.space_after = Pt(2)
        elif section and section not in PROSE_SECTIONS:
            paragraph = document.add_paragraph()
            run = paragraph.add_run(line)
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(10)
            paragraph.paragraph_format.space_before = Pt(3)
            paragraph.paragraph_format.space_after = Pt(1)
        else:
            document.add_paragraph(line)
        nonempty_seen += 1

    document.core_properties.title = title
    document.core_properties.author = "My_AGENT - Satyam Singh Bhadoriya"
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def build_cv_xlsx(
    text: str,
    title: str = "CV Reference",
    photo_bytes: bytes | None = None,
) -> bytes:
    """Create an Excel reference copy; XLSX is not recommended for ATS upload."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CV"
    sheet.column_dimensions["A"].width = 92
    sheet.column_dimensions["B"].width = 18
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"

    if photo_bytes:
        xl_image = XLImage(prepare_profile_photo(photo_bytes))
        xl_image.width = 92
        xl_image.height = 114
        sheet.add_image(xl_image, "B1")

    row = 1
    nonempty_seen = 0
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            row += 1
            continue
        upper = line.upper().rstrip(":")
        cell = sheet.cell(row=row, column=1, value=line)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if nonempty_seen == 0:
            cell.font = Font(name="Arial", size=17, bold=True, color="28306B")
            sheet.row_dimensions[row].height = 27
        elif nonempty_seen == 1 and line == upper and len(line.split()) <= 6:
            cell.font = Font(name="Arial", size=11, bold=True, color="5A4B9C")
            sheet.row_dimensions[row].height = 20
        elif upper in SECTION_HEADINGS or upper == "JD-MATCHED SKILLS":
            cell.value = upper
            cell.font = Font(name="Arial", size=11, bold=True, color="28306B")
            cell.fill = PatternFill("solid", fgColor="EEF0FA")
            sheet.row_dimensions[row].height = 22
        else:
            cell.font = Font(name="Arial", size=10, color="2A2E3A")
            sheet.row_dimensions[row].height = 18
        nonempty_seen += 1
        row += 1

    sheet["C1"] = title
    sheet.column_dimensions["C"].hidden = True
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def safe_output_stem(meta: dict) -> str:
    original = Path(meta.get("original_name", "CV.pdf")).stem
    base = re.sub(r"[^A-Za-z0-9]+", "_", original).strip("_") or "Tailored_CV"
    return f"{base}_JD_ATS"


def safe_output_name(meta: dict) -> str:
    return safe_output_stem(meta) + ".pdf"


def save_generated(pdf_bytes: bytes, filename: str) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUTPUT_DIR / filename
    path.write_bytes(pdf_bytes)
    return path
