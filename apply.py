"""Build an ATS resume PDF from a job description, straight from the terminal."""
from __future__ import annotations

import argparse
import re
from datetime import date
from pathlib import Path

from scripts.txt_to_ats_pdf import txt_to_pdf

ROOT = Path(__file__).resolve().parent
MASTER = ROOT / "applications" / "Satyam_Singh_Bhadoriya_DevOps_ATS.txt"
OUT_DIR = ROOT / "applications"
TRACKER = ROOT / "applications" / "tracker.md"

KNOWN = [
    "GitHub Actions",
    "Infrastructure as Code",
    "CloudWatch Logs",
    "Security Groups",
    "object-oriented",
    "disaster recovery",
    "pull requests",
    "Cloudflare",
    "Hostinger",
    "Kubernetes",
    "Prometheus",
    "CloudWatch",
    "Terraform",
    "Jenkins",
    "DynamoDB",
    "Datadog",
    "Grafana",
    "Ansible",
    "Hibernate",
    "Route53",
    "Python",
    "Docker",
    "Lambda",
    "Tomcat",
    "Ubuntu",
    "CentOS",
    "Apache",
    "MariaDB",
    "Windows",
    "MySQL",
    "Spring",
    "Nginx",
    "Azure",
    "Linux",
    "Helm",
    "Java",
    "Bash",
    "Shell",
    "Alloy",
    "Tempo",
    "CI/CD",
    "AWS",
    "EC2",
    "IAM",
    "S3",
    "RDS",
    "SQS",
    "SNS",
    "KMS",
    "VPC",
    "VPN",
    "CDN",
    "Git",
    "IaC",
    "OOP",
]


def slug(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", text.strip()).strip("_")
    return cleaned[:40] or "Job"


def matches_in(jd: str) -> list[str]:
    found = []
    for skill in KNOWN:
        pattern = r"(?<![A-Za-z0-9])" + re.escape(skill) + r"(?![A-Za-z0-9])"
        if re.search(pattern, jd, flags=re.IGNORECASE):
            found.append(skill)
    return found


def reorder_skill_line(line: str, matched: list[str]) -> str:
    if ":" not in line:
        return line
    label, rest = line.split(":", 1)
    parts = [p.strip() for p in rest.split(",") if p.strip()]
    hit = [p for p in parts if any(m.lower() == p.lower() or m.lower() in p.lower() for m in matched)]
    rest_parts = [p for p in parts if p not in hit]
    return f"{label}: {', '.join(hit + rest_parts)}"


def tailor(master: str, jd: str, role: str) -> str:
    matched = matches_in(jd)
    lines = master.splitlines()
    if role and len(lines) > 1:
        lines[1] = role.strip()
    out: list[str] = []
    in_skills = False
    for line in lines:
        if line.strip().upper() == "SKILLS":
            in_skills = True
            out.append(line)
            if matched:
                out.append("JD match: " + ", ".join(matched))
            continue
        if line.strip().upper() in {
            "PROFESSIONAL EXPERIENCE",
            "EDUCATION",
            "CERTIFICATIONS",
            "SUMMARY",
        }:
            in_skills = False
        if in_skills and line.strip() and ":" in line:
            out.append(reorder_skill_line(line, matched))
            continue
        out.append(line)
    return "\n".join(out) + "\n"


def guess_role(jd: str) -> str:
    for line in jd.strip().splitlines():
        if re.match(r"^\s*(job\s*title|role|position)\s*:", line, flags=re.IGNORECASE):
            return re.sub(r"^[^:]*:", "", line).strip()[:80] or "DevOps Engineer"
    first = (jd.strip().splitlines() or ["DevOps Engineer"])[0].strip()
    # A real title is short; a pasted sentence is not.
    if not first or len(first.split()) > 6:
        return "DevOps Engineer"
    return first[:80]


def guess_company(jd: str) -> str:
    for line in jd.strip().splitlines():
        if re.match(r"^\s*(company|organisation|organization|employer)\s*:", line, flags=re.IGNORECASE):
            return re.sub(r"^[^:]*:", "", line).strip()[:60] or "Company"
    return "Company"


END_MARKERS = {"END", "DONE", "BAS", "OK"}


def read_jd_from_terminal() -> str:
    print("=" * 60)
    print("  GIVE ME YOUR JD  (job description yahan paste karo)")
    print("=" * 60)
    print("Paste karne ke baad nayi line pe 'END' likh ke Enter dabao.\n")
    lines: list[str] = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip().upper() in END_MARKERS:
            break
        lines.append(line)
    return "\n".join(lines).strip()


def ask(label: str, default: str) -> str:
    answer = input(f"{label} [{default}]: ").strip()
    if not answer or answer.upper() in END_MARKERS:
        return default
    return answer


def append_tracker(company: str, role: str, pdf: Path) -> None:
    row = f"| {date.today().isoformat()} | {company} | {role} | local CLI | - | drafted | {pdf.name} |\n"
    if TRACKER.exists():
        text = TRACKER.read_text(encoding="utf-8")
        if pdf.name in text:
            return
        TRACKER.write_text(text.rstrip() + "\n" + row, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build an ATS CV PDF from a JD.")
    parser.add_argument("--jd", help="JD wali .txt file. Na do to terminal JD maangega.")
    parser.add_argument("--company", default="", help="Company name (file name ke liye)")
    parser.add_argument("--role", default="", help="Job title; blank ho to JD se uthayega")
    args = parser.parse_args()

    if not MASTER.is_file():
        raise SystemExit(f"Master CV nahi mili: {MASTER}")

    if args.jd:
        jd_path = Path(args.jd).expanduser().resolve()
        if not jd_path.is_file():
            raise SystemExit(f"JD file nahi mili: {jd_path}")
        jd = jd_path.read_text(encoding="utf-8")
        company = args.company.strip() or guess_company(jd)
        role = args.role.strip() or guess_role(jd)
    else:
        jd = read_jd_from_terminal()
        if not jd:
            raise SystemExit("JD khali tha. Dubara chalao aur JD paste karo.")
        print()
        role = args.role.strip() or ask("Job title", guess_role(jd))
        company = args.company.strip() or ask("Company", guess_company(jd))
        print()

    master = MASTER.read_text(encoding="utf-8")
    tailored = tailor(master, jd, role)

    stem = f"Satyam_Singh_Bhadoriya_{slug(role)}_{slug(company)}_ATS"
    txt_path = OUT_DIR / f"{stem}.txt"
    pdf_path = OUT_DIR / f"{stem}.pdf"
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    txt_path.write_text(tailored, encoding="utf-8")
    txt_to_pdf(txt_path, pdf_path)
    append_tracker(company, role, pdf_path)

    print("Matched skills:", ", ".join(matches_in(jd)) or "(JD mein known skills kam mili)")
    print("TXT:", txt_path)
    print("PDF:", pdf_path)
    print("\nCV ready. Yeh PDF upload karo.")


if __name__ == "__main__":
    main()
