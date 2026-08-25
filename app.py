"""Standalone browser UI for creating truthful ATS CVs from a JD."""
from __future__ import annotations

import base64
from pathlib import Path

import streamlit as st

from cv_engine import (
    build_ats_pdf,
    build_ats_docx,
    build_cv_xlsx,
    build_cold_email,
    build_tailored_text,
    compose_cv_text,
    extract_pdf_text,
    load_existing_cv,
    match_jd,
    random_output_stem,
    save_existing_cv,
    save_generated,
)

APP_PIN = "6932"
ROOT_DIR = Path(__file__).resolve().parent
ASSETS_DIR = ROOT_DIR / "assets"
OWNER_PHOTO = ASSETS_DIR / "owner_photo.jpg"
OWNER_CV = ROOT_DIR / "profile" / "Satyam_Dev_Resume_ATS.pdf"

# Owner details ship with the code so every copy shows the same profile.
OWNER_PROFILE = {
    "name": "Satyam Singh Bhadoriya",
    "role": "DevOps Engineer",
    "email": "Satyam.bhadoriya6932@gmail.com",
    "phone": "8962373424",
    "location": "Indore, Madhya Pradesh",
    "education": "Master of Computer Application — RGPV Bhopal (2023 - 2025)",
    "certification": "DevOps Engineer Foundation Course — Skilling Academy (2025)",
    "core_skills": (
        "AWS, Linux, Git, Docker, Kubernetes, Terraform, Jenkins, CI/CD, "
        "Infrastructure as Code"
    ),
}

st.set_page_config(
    page_title="My_AGENT - ATS CV Builder",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="expanded",
)

THEMES = {
    "Dark": {
        "bg": "#141726",
        "glow": "radial-gradient(1100px 520px at 12% -10%, #4b4a8f1f, transparent),"
               " radial-gradient(900px 460px at 88% 4%, #3f5a8f1f, transparent)",
        "panel": "#1c2033",
        "panel_soft": "#22273d",
        "border": "#2f3550",
        "text": "#dfe3f2",
        "muted": "#9298b4",
        "accent": "#6c6fbf",
        "accent2": "#5a7fb8",
        "hero": "linear-gradient(135deg, #2b2b52 0%, #33395f 55%, #33496b 100%)",
        "shadow": "0 12px 28px rgba(8,10,20,.32)",
    },
    "Light": {
        "bg": "#f2f2f8",
        "glow": "radial-gradient(1100px 520px at 12% -10%, #6c6fbf14, transparent),"
               " radial-gradient(900px 460px at 88% 4%, #5a7fb814, transparent)",
        "panel": "#ffffff",
        "panel_soft": "#f7f7fc",
        "border": "#dcdcea",
        "text": "#232741",
        "muted": "#616784",
        "accent": "#5f62b0",
        "accent2": "#4f74ab",
        "hero": "linear-gradient(135deg, #43457d 0%, #46527f 55%, #43607f 100%)",
        "shadow": "0 10px 24px rgba(60,64,110,.10)",
    },
}

st.session_state.setdefault("theme", "Dark")
st.session_state.setdefault("screen", "home")
st.session_state.setdefault("change_cv", False)


def paint(theme_name: str) -> None:
    t = THEMES[theme_name]
    st.markdown(
        f"""
        <style>
        .stApp {{ background: {t["bg"]}; background-image: {t["glow"]}; }}
        .block-container {{ padding-top: 2.2rem; max-width: 900px; }}
        header[data-testid="stHeader"] {{ background: transparent; }}
        html, body, .stApp, p, span, label, li, div[data-testid="stMarkdownContainer"] {{
            color: {t["text"]};
        }}
        h1, h2, h3, h4 {{ color: {t["text"]}; letter-spacing: -.02em; }}

        .hero {{
            padding: 2.4rem 1.6rem; border-radius: 22px; text-align: center;
            background: {t["hero"]}; color: #f8fbff; margin-bottom: 1.1rem;
            box-shadow: {t["shadow"]}; border: 1px solid {t["border"]};
        }}
        .hero h1, .hero h1 * {{
            margin: 0; font-size: 2.4rem; font-weight: 750;
            color: #ffffff !important;
            text-shadow: 0 1px 6px rgba(0,0,0,.22);
        }}
        .hero .by {{
            margin-top: .5rem; font-size: .95rem; letter-spacing: .14em;
            text-transform: uppercase; color: #cfe6ff;
        }}
        .hero .flow {{
            margin-top: 1rem; display: inline-block; padding: .5rem 1rem;
            border-radius: 999px; background: rgba(255,255,255,.12);
            border: 1px solid rgba(255,255,255,.22); font-size: .9rem;
            color: #eaf4ff;
        }}
        .card {{
            background: {t["panel"]}; border: 1px solid {t["border"]};
            border-radius: 16px; padding: 1.15rem 1.25rem; margin: .55rem 0;
            box-shadow: {t["shadow"]};
        }}
        .card h4 {{ margin: 0 0 .35rem; font-size: 1.05rem; }}
        .card p {{ margin: 0; color: {t["muted"]}; font-size: .92rem; }}
        .step {{
            display: inline-flex; align-items: center; gap: .55rem;
            font-weight: 700; font-size: 1.15rem; margin: .2rem 0 .6rem;
        }}
        .step .num {{
            width: 1.85rem; height: 1.85rem; border-radius: 50%;
            display: inline-flex; align-items: center; justify-content: center;
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
            color: #fff; font-size: .95rem;
        }}
        .pill {{
            display: inline-block; padding: .3rem .7rem; margin: .18rem .25rem 0 0;
            border-radius: 999px; font-size: .82rem; background: {t["panel_soft"]};
            border: 1px solid {t["border"]}; color: {t["text"]};
        }}
        div.stButton > button, div.stDownloadButton > button {{
            width: 100%; min-height: 3.1rem; border-radius: 14px;
            font-weight: 700; font-size: 1rem; letter-spacing: .01em;
            background: {t["panel"]}; color: {t["text"]};
            border: 1.5px solid {t["border"]};
            transition: transform .16s ease, box-shadow .16s ease,
                        background .16s ease, border-color .16s ease, color .16s ease;
        }}
        div.stButton > button:hover, div.stDownloadButton > button:hover {{
            transform: translateY(-2px);
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
            color: #ffffff !important; border-color: transparent;
            box-shadow: 0 8px 18px {t["accent"]}33;
        }}
        div.stButton > button:hover p, div.stDownloadButton > button:hover p {{
            color: #ffffff !important;
        }}
        div.stButton > button:active, div.stDownloadButton > button:active {{
            transform: translateY(-1px) scale(.99);
        }}
        div.stButton > button[kind="primary"],
        div.stDownloadButton > button[kind="primary"] {{
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
            color: #fff; border: none;
            box-shadow: 0 6px 16px {t["accent"]}2e;
        }}
        div.stButton > button[kind="primary"]:hover,
        div.stDownloadButton > button[kind="primary"]:hover {{
            background: linear-gradient(135deg, {t["accent2"]}, {t["accent"]});
            box-shadow: 0 10px 22px {t["accent2"]}3d;
        }}
        .option-card {{
            background: {t["panel"]}; border: 1.5px solid {t["border"]};
            border-radius: 18px; padding: 1.35rem 1.25rem 1.1rem;
            margin: .55rem 0 .3rem; box-shadow: {t["shadow"]};
            position: relative; overflow: hidden; transition: .2s ease;
        }}
        .option-card::before {{
            content: ""; position: absolute; inset: 0 0 auto 0; height: 4px;
            background: linear-gradient(90deg, {t["accent"]}, {t["accent2"]});
        }}
        .option-card:hover {{
            transform: translateY(-3px); border-color: {t["accent"]};
        }}
        .option-card .ic {{
            width: 3rem; height: 3rem; border-radius: 14px; font-size: 1.4rem;
            display: flex; align-items: center; justify-content: center;
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
            color: #fff; margin-bottom: .7rem;
        }}
        .option-card h4 {{ margin: 0 0 .3rem; font-size: 1.12rem; }}
        .option-card p {{ margin: 0; color: {t["muted"]}; font-size: .9rem; }}
        .st-key-theme_toggle div.stButton > button {{
            width: 3.1rem; min-height: 3.1rem; padding: 0;
            border-radius: 50%; font-size: 1.35rem;
            background: {t["panel"]}; border: 1.5px solid {t["border"]};
            box-shadow: {t["shadow"]};
        }}
        .st-key-theme_toggle div.stButton > button:hover {{
            transform: rotate(-18deg) scale(1.08);
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
        }}
        div[data-testid="stFileUploader"] section,
        div[data-testid="stFileUploaderDropzone"] {{
            background: {t["panel_soft"]}; border: 1.5px dashed {t["border"]};
            border-radius: 14px;
        }}
        .stTextArea textarea {{
            background: {t["panel"]}; color: {t["text"]};
            border: 1px solid {t["border"]}; border-radius: 14px;
            font-size: .95rem;
        }}
        div[data-testid="stMetric"] {{
            background: {t["panel"]}; border: 1px solid {t["border"]};
            border-radius: 14px; padding: .85rem 1rem;
        }}
        div[data-testid="stMetricValue"] {{ color: {t["text"]}; }}
        hr {{ border-color: {t["border"]}; }}
        section[data-testid="stSidebar"] {{
            background: {t["panel"]};
            border-right: 1px solid {t["border"]};
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {{
            padding-top: 1.2rem;
        }}
        .side-brand {{
            padding: .9rem .25rem 1.1rem; text-align: center;
            font-weight: 800; font-size: 1.15rem; letter-spacing: .04em;
            color: {t["text"]}; animation: agentFadeUp .45s ease both;
        }}
        .passport-photo {{
            width: 100%; max-width: 165px; aspect-ratio: 35 / 45;
            object-fit: cover; object-position: center top;
            border-radius: 14px; display: block;
            border: 2px solid {t["accent"]};
            box-shadow: {t["shadow"]};
            animation: agentFadeUp .5s ease both;
        }}
        .profile-card p {{ margin: .18rem 0; font-size: .9rem; }}
        .profile-card .profile-role {{
            color: {t["accent"]}; font-weight: 700; font-size: .95rem;
            margin-bottom: .5rem;
        }}
        .side-brand span {{
            display: block; margin-top: .2rem; font-weight: 500;
            font-size: .72rem; letter-spacing: .12em; color: {t["muted"]};
            text-transform: uppercase;
        }}

        .privacy {{
            text-align: center; color: {t["muted"]}; font-size: .86rem;
            line-height: 1.55; max-width: 640px; margin: 0 auto .4rem;
            animation: agentFadeUp .6s ease both;
        }}
        .privacy strong {{ color: {t["text"]}; font-weight: 650; }}

        @keyframes agentFadeUp {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to   {{ opacity: 1; transform: translateY(0); }}
        }}
        @keyframes agentPulse {{
            0%, 100% {{ box-shadow: 0 0 0 0 {t["accent"]}00; }}
            50%      {{ box-shadow: 0 0 0 6px {t["accent"]}1f; }}
        }}
        @keyframes agentBarGrow {{
            from {{ transform: scaleX(0); }}
            to   {{ transform: scaleX(1); }}
        }}

        .hero {{ animation: agentFadeUp .55s ease both; }}
        .step {{ animation: agentFadeUp .45s ease both; }}
        .step .num {{ animation: agentPulse 3.2s ease-in-out infinite; }}
        .card {{
            animation: agentFadeUp .5s ease both;
            transition: transform .18s ease, box-shadow .18s ease,
                        border-color .18s ease;
        }}
        .card:hover {{ transform: translateY(-2px); border-color: {t["accent"]}; }}
        .option-card {{ animation: agentFadeUp .5s ease both; }}
        .option-card.d2 {{ animation-delay: .09s; }}
        .option-card.d3 {{ animation-delay: .18s; }}
        .option-card::before {{
            transform-origin: left center;
            animation: agentBarGrow .7s ease both;
        }}
        .option-card .ic {{ transition: transform .25s ease; }}
        .option-card:hover .ic {{ transform: translateY(-2px) scale(1.06); }}
        .pill {{ transition: transform .16s ease, border-color .16s ease; }}
        .pill:hover {{ transform: translateY(-2px); border-color: {t["accent"]}; }}
        div[data-testid="stExpander"] {{ animation: agentFadeUp .5s ease both; }}
        div[data-testid="stMetric"] {{
            animation: agentFadeUp .5s ease both;
            transition: transform .18s ease, border-color .18s ease;
        }}
        div[data-testid="stMetric"]:hover {{
            transform: translateY(-2px); border-color: {t["accent"]};
        }}
        div[data-testid="stAlert"] {{ animation: agentFadeUp .45s ease both; }}
        .stTextInput input, .stTextArea textarea {{
            transition: border-color .18s ease, box-shadow .18s ease;
        }}
        .stTextInput input:focus, .stTextArea textarea:focus {{
            border-color: {t["accent"]} !important;
            box-shadow: 0 0 0 3px {t["accent"]}26;
        }}
        @media (prefers-reduced-motion: reduce) {{
            .hero, .step, .step .num, .card, .option-card,
            .option-card::before, .privacy, div[data-testid="stExpander"],
            div[data-testid="stMetric"], div[data-testid="stAlert"] {{
                animation: none;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def backdrop_photo_uri() -> str:
    """Optional user photo backdrop: drop a file in assets/background.<ext>."""
    for ext in ("jpg", "jpeg", "png", "webp"):
        path = ASSETS_DIR / f"background.{ext}"
        if path.exists():
            mime = "jpeg" if ext in {"jpg", "jpeg"} else ext
            encoded = base64.b64encode(path.read_bytes()).decode("ascii")
            return f"data:image/{mime};base64,{encoded}"
    return ""


def owner_photo_uri() -> str:
    """Base64 data URI for the bundled passport-size owner photo."""
    if not OWNER_PHOTO.exists():
        return ""
    encoded = base64.b64encode(OWNER_PHOTO.read_bytes()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def photo_backdrop(theme_name: str) -> None:
    """Slow Ken Burns zoom on the user's own background image, if present."""
    uri = backdrop_photo_uri()
    if not uri:
        return
    opacity = ".22" if theme_name == "Dark" else ".16"
    st.markdown(
        f"""
        <style>
        .bg-photo {{
            position: fixed; inset: 0; z-index: -1; pointer-events: none;
            background-image: url("{uri}");
            background-size: cover; background-position: center;
            opacity: {opacity}; filter: blur(1.5px) saturate(.9);
            animation: agentZoom 34s ease-in-out infinite alternate;
        }}
        @keyframes agentZoom {{
            0%   {{ transform: scale(1) translate3d(0, 0, 0); }}
            100% {{ transform: scale(1.09) translate3d(-1.5%, -1.5%, 0); }}
        }}
        @media (prefers-reduced-motion: reduce) {{
            .bg-photo {{ animation: none; }}
        }}
        </style>
        <div class="bg-photo"></div>
        """,
        unsafe_allow_html=True,
    )


def animated_backdrop(theme_name: str) -> None:
    """Slow moving gradient + grid backdrop for the front page (CSS only)."""
    t = THEMES[theme_name]
    blob = "26" if theme_name == "Dark" else "1c"
    grid = "0.05" if theme_name == "Dark" else "0.035"
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image:
                radial-gradient(620px 420px at 18% 12%, {t["accent"]}{blob}, transparent 70%),
                radial-gradient(560px 400px at 82% 18%, {t["accent2"]}{blob}, transparent 70%),
                radial-gradient(520px 380px at 50% 92%, {t["accent"]}{blob}, transparent 70%);
            background-repeat: no-repeat;
            background-size: 140% 140%, 130% 130%, 150% 150%;
            animation: agentDrift 28s ease-in-out infinite alternate;
        }}
        .stApp::before {{
            content: ""; position: fixed; inset: -20% -20% -20% -20%;
            pointer-events: none; z-index: 0;
            background-image:
                linear-gradient(rgba(125,140,220,{grid}) 1px, transparent 1px),
                linear-gradient(90deg, rgba(125,140,220,{grid}) 1px, transparent 1px);
            background-size: 62px 62px, 62px 62px;
            animation: agentGrid 42s linear infinite;
        }}
        .block-container {{ position: relative; z-index: 1; }}
        .hero {{ position: relative; overflow: hidden; }}
        .hero::after {{
            content: ""; position: absolute; top: -60%; left: -35%;
            width: 45%; height: 220%; pointer-events: none;
            background: linear-gradient(
                100deg, transparent, rgba(255,255,255,.16), transparent);
            transform: rotate(14deg);
            animation: agentSheen 7.5s ease-in-out infinite;
        }}
        @keyframes agentDrift {{
            0%   {{ background-position: 0% 0%, 100% 0%, 50% 100%; }}
            50%  {{ background-position: 30% 20%, 70% 25%, 40% 75%; }}
            100% {{ background-position: 12% 40%, 88% 8%, 60% 90%; }}
        }}
        @keyframes agentGrid {{
            0%   {{ transform: translate3d(0, 0, 0); }}
            100% {{ transform: translate3d(62px, 62px, 0); }}
        }}
        @keyframes agentSheen {{
            0%   {{ left: -35%; opacity: 0; }}
            18%  {{ opacity: 1; }}
            60%  {{ left: 115%; opacity: 0; }}
            100% {{ left: 115%; opacity: 0; }}
        }}
        @media (prefers-reduced-motion: reduce) {{
            .stApp, .stApp::before, .hero::after {{ animation: none; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def go(screen: str) -> None:
    st.session_state["screen"] = screen
    st.rerun()


def activate_cv(meta: dict, text: str, pdf_bytes: bytes) -> None:
    st.session_state["cv_meta"] = meta
    st.session_state["cv_text"] = text
    st.session_state["cv_pdf"] = pdf_bytes
    st.session_state["cv_ready"] = True
    st.session_state.pop("generated", None)


def clear_cv_session() -> None:
    for key in ("cv_meta", "cv_text", "cv_pdf", "cv_ready", "generated"):
        st.session_state.pop(key, None)
    st.session_state["change_cv"] = False


def step(number: int, label: str) -> None:
    st.markdown(
        f'<div class="step"><span class="num">{number}</span>{label}</div>',
        unsafe_allow_html=True,
    )


def exit_button(key: str) -> None:
    if st.button("✕  Exit to Front Page", key=key):
        clear_cv_session()
        go("home")


@st.dialog("Enter PIN to continue")
def pin_dialog() -> None:
    st.write("My_AGENT is locked. Enter your PIN to continue.")
    entered = st.text_input("PIN", type="password", max_chars=8, key="pin_input")
    if st.button("Unlock", type="primary", use_container_width=True):
        if entered.strip() == APP_PIN:
            st.session_state["unlocked"] = True
            st.rerun()
        else:
            st.error("Incorrect PIN. Please try again.")


def build_all_formats(
    text: str,
    stem: str | None = None,
    photo_bytes: bytes | None = None,
) -> dict:
    stem = stem or random_output_stem()
    pdf_bytes = build_ats_pdf(text, title=stem, photo_bytes=photo_bytes)
    docx_bytes = build_ats_docx(text, title=stem, photo_bytes=photo_bytes)
    xlsx_bytes = build_cv_xlsx(text, title=stem, photo_bytes=photo_bytes)
    return {
        "pdf": {
            "label": "PDF",
            "data": pdf_bytes,
            "filename": stem + ".pdf",
            "mime": "application/pdf",
            "path": str(save_generated(pdf_bytes, stem + ".pdf")),
        },
        "docx": {
            "label": "DOCX",
            "data": docx_bytes,
            "filename": stem + ".docx",
            "mime": (
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            "path": str(save_generated(docx_bytes, stem + ".docx")),
        },
        "xlsx": {
            "label": "XLSX",
            "data": xlsx_bytes,
            "filename": stem + ".xlsx",
            "mime": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "path": str(save_generated(xlsx_bytes, stem + ".xlsx")),
        },
    }


@st.dialog("Download your CV")
def download_format_dialog(generated: dict) -> None:
    st.write("Which format would you like to download?")
    st.success("PDF and DOCX are both ATS-friendly.")
    for key, description in (
        ("pdf", "PDF — selectable text, clean one-column layout (Recommended)"),
        ("docx", "DOCX — editable Word file, ATS-friendly"),
    ):
        file_info = generated["files"][key]
        st.markdown(f"**{description}**")
        st.download_button(
            f"⬇ Download {file_info['label']}",
            data=file_info["data"],
            file_name=file_info["filename"],
            mime=file_info["mime"],
            type="primary" if key == "pdf" else "secondary",
            use_container_width=True,
            key=f"download_{key}",
        )
    st.warning(
        "XLSX — Excel reference copy only. Do not upload it to a job portal or "
        "ATS; use the PDF or DOCX when you apply."
    )
    excel_info = generated["files"]["xlsx"]
    st.download_button(
        "⬇ Download XLSX (Reference only)",
        data=excel_info["data"],
        file_name=excel_info["filename"],
        mime=excel_info["mime"],
        use_container_width=True,
        key="download_xlsx",
    )


paint(st.session_state["theme"])

if not st.session_state.get("unlocked"):
    animated_backdrop(st.session_state["theme"])
    photo_backdrop(st.session_state["theme"])
    st.markdown(
        """
        <div class="hero">
          <h1>Welcome to My_AGENT</h1>
          <div class="by">Created by Satyam Singh Bhadoriya</div>
          <div class="flow">🔒 Locked — enter your PIN to unlock</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    pin_dialog()
    st.stop()

with st.sidebar:
    st.markdown(
        '<div class="side-brand">My_AGENT<span>Career Toolkit</span></div>',
        unsafe_allow_html=True,
    )
    if st.button("👤  My Profile", use_container_width=True, key="side_profile"):
        clear_cv_session()
        go("profile")
    if st.button(
        "✉  Cold Mail for Referral",
        use_container_width=True,
        key="side_cold_mail",
    ):
        clear_cv_session()
        go("cold_mail")
    st.divider()
    if st.button("⚙  Settings", use_container_width=True, key="side_settings"):
        clear_cv_session()
        go("settings")

is_dark = st.session_state["theme"] == "Dark"
spacer, toggle_col = st.columns([8, 1])
with toggle_col:
    with st.container(key="theme_toggle"):
        if st.button(
            "☀" if is_dark else "🌙",
            help="Switch to Light mode" if is_dark else "Switch to Dark mode",
        ):
            st.session_state["theme"] = "Light" if is_dark else "Dark"
            st.rerun()

screen = st.session_state["screen"]
if screen == "home":
    animated_backdrop(st.session_state["theme"])
    photo_backdrop(st.session_state["theme"])

hero_flow = (
    "Upload CV → Analyse profile → Create HR email"
    if screen == "cold_mail"
    else "Upload CV → Give JD → Generate ATS CV → Download"
)
st.markdown(
    f"""
    <div class="hero">
      <h1>Welcome to My_AGENT</h1>
      <div class="by">Created by Satyam Singh Bhadoriya</div>
      <div class="flow">{hero_flow}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

existing = load_existing_cv()

if screen == "home":
    step(1, "Choose an option")
    left, right, third = st.columns(3)
    with left:
        st.markdown(
            '<div class="option-card d1"><div class="ic">⬆</div>'
            "<h4>1. Upload CV</h4>"
            "<p>Upload a new CV PDF. Your details are saved on this computer.</p>"
            "</div>",
            unsafe_allow_html=True,
        )
        if st.button("⬆  Upload CV", type="primary", key="home_upload"):
            clear_cv_session()
            go("upload")
    with right:
        saved_name = (
            existing[0].get("original_name", "CV.pdf")
            if existing
            else "No CV saved yet"
        )
        st.markdown(
            f'<div class="option-card d2"><div class="ic">📁</div>'
            f"<h4>2. Existing CV</h4><p>{saved_name}</p></div>",
            unsafe_allow_html=True,
        )
        if st.button("📁  Existing CV", type="primary", key="home_existing"):
            clear_cv_session()
            go("existing")
    with third:
        st.markdown(
            '<div class="option-card d3"><div class="ic">✍</div>'
            "<h4>3. Make your CV</h4>"
            "<p>Fill in your details and build a new ATS CV.</p></div>",
            unsafe_allow_html=True,
        )
        if st.button("✍  Make your CV", type="primary", key="home_make"):
            clear_cv_session()
            st.session_state.pop("manual_generated", None)
            go("make")

elif screen == "upload":
    step(1, "Upload your CV")
    uploaded = st.file_uploader(
        "CV (PDF)",
        type=["pdf"],
        help="Upload a selectable-text PDF. Scanned image PDFs are not supported.",
    )
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        try:
            extracted = extract_pdf_text(file_bytes)
            st.success(f"Text extracted: {len(extracted)} characters")
            if st.button("Use this CV", type="primary"):
                meta = save_existing_cv(file_bytes, uploaded.name, extracted)
                activate_cv(meta, extracted, file_bytes)
                st.rerun()
        except Exception as exc:
            st.error(str(exc))
    exit_button("exit_upload")

elif screen == "existing":
    step(1, "Existing CV")
    if existing is None:
        st.warning(
            "No CV is saved yet. Go back to the front page and choose **Upload CV**."
        )
    else:
        meta, text, pdf_bytes = existing
        st.markdown(
            f"""
            <div class="card">
              <h4>{meta.get("original_name", "CV.pdf")}</h4>
              <p>Saved: {meta.get("uploaded_at", "-")} &nbsp;·&nbsp;
              Extracted characters: {meta.get("characters", len(text))}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Continue", type="primary"):
                activate_cv(meta, text, pdf_bytes)
                st.session_state["change_cv"] = False
                st.rerun()
        with col_b:
            if st.button("Change CV"):
                clear_cv_session()
                st.session_state["change_cv"] = True
                st.rerun()

        if st.session_state.get("change_cv"):
            replacement = st.file_uploader(
                "New CV (PDF)", type=["pdf"], key="replacement_cv"
            )
            if replacement is not None:
                try:
                    replacement_bytes = replacement.getvalue()
                    replacement_text = extract_pdf_text(replacement_bytes)
                    if st.button("Save and continue", type="primary"):
                        new_meta = save_existing_cv(
                            replacement_bytes, replacement.name, replacement_text
                        )
                        activate_cv(new_meta, replacement_text, replacement_bytes)
                        st.session_state["change_cv"] = False
                        st.rerun()
                except Exception as exc:
                    st.error(str(exc))
    exit_button("exit_existing")

elif screen == "profile":
    step(1, "My Profile")
    photo_col, detail_col = st.columns([1, 2.4])
    with photo_col:
        photo_uri = owner_photo_uri()
        if photo_uri:
            st.markdown(
                f'<img class="passport-photo" src="{photo_uri}" '
                'alt="Satyam Singh Bhadoriya" />',
                unsafe_allow_html=True,
            )
        else:
            st.info("Profile photo not found in the assets folder.")
    with detail_col:
        st.markdown(
            f"""
            <div class="card profile-card">
              <h4>{OWNER_PROFILE["name"]}</h4>
              <p class="profile-role">{OWNER_PROFILE["role"]}</p>
              <p><strong>Email:</strong> {OWNER_PROFILE["email"]}</p>
              <p><strong>Phone:</strong> {OWNER_PROFILE["phone"]}</p>
              <p><strong>Location:</strong> {OWNER_PROFILE["location"]}</p>
              <p><strong>Education:</strong> {OWNER_PROFILE["education"]}</p>
              <p><strong>Certification:</strong> {OWNER_PROFILE["certification"]}</p>
              <p><strong>Core skills:</strong> {OWNER_PROFILE["core_skills"]}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if OWNER_CV.exists():
        st.download_button(
            "⬇  My CV for Reference",
            data=OWNER_CV.read_bytes(),
            file_name="Satyam_Singh_Bhadoriya_CV.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
            key="owner_cv_download",
        )
        st.caption(
            "This is the reference resume whose layout and section order every "
            "generated CV follows."
        )
    else:
        st.warning("Reference CV file is missing from the profile folder.")
    exit_button("exit_profile")

elif screen == "cold_mail":
    step(1, "Cold Mail for Referral")
    st.caption(
        "Upload a selectable-text CV. My_AGENT will identify whether the "
        "candidate is a fresher, has internship exposure, or has professional "
        "experience, then prepare a concise email using only CV facts."
    )
    cold_cv = st.file_uploader(
        "Upload CV (PDF)",
        type=["pdf"],
        key="cold_mail_cv",
        help="The PDF must contain selectable text. Scanned image PDFs are not supported.",
    )
    if cold_cv is not None:
        try:
            cold_text = extract_pdf_text(cold_cv.getvalue())
            st.success(f"CV analysed successfully: {len(cold_text)} characters read.")

            target_col, company_col = st.columns(2)
            with target_col:
                target_role = st.text_input(
                    "Target role (optional)",
                    key="cold_target_role",
                    placeholder="e.g. Java Developer",
                )
            with company_col:
                target_company = st.text_input(
                    "Company name (optional)",
                    key="cold_company",
                    placeholder="e.g. Acme Technologies",
                )

            cold_mail = build_cold_email(
                cold_text,
                target_role=target_role,
                company=target_company,
            )
            mail_key = str(
                abs(
                    hash(
                        (
                            cold_cv.name,
                            len(cold_cv.getvalue()),
                            target_role,
                            target_company,
                        )
                    )
                )
            )

            st.divider()
            step(2, "Your HR email is ready")
            info_col, skill_col = st.columns(2)
            info_col.metric("Candidate profile", cold_mail["status"])
            skill_col.metric("Verified skills used", len(cold_mail["skills"]))

            if cold_mail["skills"]:
                st.markdown("**Skills verified from the CV**")
                st.markdown(
                    "".join(
                        f'<span class="pill">{skill}</span>'
                        for skill in cold_mail["skills"]
                    ),
                    unsafe_allow_html=True,
                )

            subject = st.text_input(
                "Email subject",
                value=str(cold_mail["subject"]),
                key=f"cold_subject_{mail_key}",
            )
            body = st.text_area(
                "Email body",
                value=str(cold_mail["body"]),
                height=390,
                key=f"cold_body_{mail_key}",
            )
            email_file = f"Subject: {subject}\n\n{body}\n"
            st.download_button(
                "⬇  Download Email as TXT",
                data=email_file.encode("utf-8"),
                file_name="My_AGENT_Cold_Email.txt",
                mime="text/plain",
                type="primary",
                use_container_width=True,
                key="cold_download",
            )
            st.caption(
                "Review the recipient, role, and company before sending. The "
                "generated text never adds unverified employment or skills."
            )
        except Exception as exc:
            st.error(str(exc))
    exit_button("exit_cold_mail")

elif screen == "settings":
    step(1, "Settings")
    st.markdown(
        """
        <div class="card">
          <h4>Session access</h4>
          <p>Log out to lock My_AGENT and return to the PIN screen.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(
        "↪  Log out",
        type="primary",
        use_container_width=True,
        key="settings_logout",
    ):
        current_theme = st.session_state.get("theme", "Dark")
        st.session_state.clear()
        st.session_state["theme"] = current_theme
        st.session_state["screen"] = "home"
        st.rerun()
    exit_button("exit_settings")

elif screen == "make":
    step(1, "Make your CV")
    st.caption(
        "No field is mandatory. The more details you add, the more complete "
        "your CV will be."
    )
    form_left, form_right = st.columns(2)
    with form_left:
        full_name = st.text_input("Name", key="make_name", placeholder="Your full name")
        phone = st.text_input("Number", key="make_phone", placeholder="Mobile number")
    with form_right:
        email = st.text_input("Email", key="make_email", placeholder="you@example.com")
        address = st.text_input(
            "Address",
            key="make_address",
            placeholder="House / street / locality",
        )
    city_col, state_col = st.columns(2)
    with city_col:
        city = st.text_input("City", key="make_city", placeholder="City")
    with state_col:
        state = st.text_input("State", key="make_state", placeholder="State")
    headline = st.text_input(
        "Job Title / Role",
        key="make_headline",
        placeholder="DevOps Engineer, Java Full Stack Developer, Python Developer...",
    )
    profile_photo = st.file_uploader(
        "Profile photo (optional)",
        type=["jpg", "jpeg", "png", "webp"],
        key="make_photo",
        help="If added, the photo is placed at the top-right of the generated CV.",
    )
    photo_bytes = profile_photo.getvalue() if profile_photo is not None else None
    if photo_bytes:
        st.caption(
            "This photo will appear at the top-right of the PDF and Word CV. "
            "Some ATS software ignore images, so keep the rest of the CV text-based."
        )

    st.markdown("**CV sections** — fill in whichever apply to you")
    summary = st.text_area(
        "Professional Summary",
        key="make_summary",
        height=90,
        placeholder="2–4 lines about your background and goal",
    )
    skills_text = st.text_area(
        "Skills",
        key="make_skills",
        height=80,
        placeholder="Python, Java, React, AWS...",
    )
    exp_col, proj_col = st.columns(2)
    with exp_col:
        experience = st.text_area(
            "Professional Experience",
            key="make_experience",
            height=140,
            placeholder="Company, role, dates, work done",
        )
    with proj_col:
        projects = st.text_area(
            "Projects",
            key="make_projects",
            height=140,
            placeholder="Project name, what you built, stack",
        )
    edu_col, cert_col = st.columns(2)
    with edu_col:
        education = st.text_area(
            "Education",
            key="make_education",
            height=110,
            placeholder="Degree, college, year",
        )
    with cert_col:
        certifications = st.text_area(
            "Certifications",
            key="make_certifications",
            height=110,
            placeholder="Certificate name, issuer, year",
        )
    intern_col, extra_col, lang_col = st.columns(3)
    with intern_col:
        internships = st.text_area(
            "Internships",
            key="make_internships",
            height=90,
            placeholder="Company, role, dates",
        )
    with extra_col:
        achievements = st.text_area(
            "Achievements",
            key="make_achievements",
            height=90,
            placeholder="Awards, rankings, highlights",
        )
    with lang_col:
        languages = st.text_area(
            "Languages",
            key="make_languages",
            height=90,
            placeholder="English, Hindi...",
        )

    st.markdown("**Give me your all details**")
    details = st.text_area(
        "All details",
        height=180,
        label_visibility="collapsed",
        key="make_details",
        placeholder=(
            "Extra notes, a company JD, additional skills, college or home "
            "address — anything missed in the fields above"
        ),
    )
    st.caption(
        "Note: this box accepts a company JD, extra skills, experience, college "
        "details, or your home address. Nothing here is mandatory."
    )

    st.caption(
        "Only the layout and section order come from the reference resume — the "
        "content is entirely what you enter. Leave the summary empty and a "
        "professional summary is written from your role and skills."
    )

    form_values = {
        "name": full_name,
        "email": email,
        "phone": phone,
        "address": address,
        "city": city,
        "state": state,
        "headline": headline,
        "summary": summary,
        "skills": skills_text,
        "experience": experience,
        "projects": projects,
        "education": education,
        "certifications": certifications,
        "internships": internships,
        "achievements": achievements,
        "languages": languages,
        "details": details,
        "photo": (
            f"{profile_photo.name}:{len(photo_bytes)}" if photo_bytes else ""
        ),
    }
    signature = repr(form_values)

    if st.button("⚡  Make CV", type="primary", key="make_cv_btn"):
        try:
            composed = compose_cv_text(
                name=full_name,
                email=email,
                phone=phone,
                address=address,
                city=city,
                state=state,
                headline=headline,
                details=details,
                sections={
                    "summary": summary,
                    "skills": skills_text,
                    "experience": experience,
                    "projects": projects,
                    "education": education,
                    "certifications": certifications,
                    "internships": internships,
                    "achievements": achievements,
                    "languages": languages,
                },
            )
            skill_source = "\n".join(
                [skills_text, experience, projects, details, internships]
            )
            st.session_state["manual_generated"] = {
                "files": build_all_formats(composed, photo_bytes=photo_bytes),
                "skills": match_jd(composed, skill_source or composed)["matched"],
                "text": composed,
                "signature": signature,
            }
        except Exception as exc:
            st.error(str(exc))

    manual = st.session_state.get("manual_generated")
    if manual:
        st.divider()
        step(2, "Your CV is ready")
        if manual.get("signature") != signature:
            st.warning(
                "Your form details have changed. The CV below uses the older "
                "details — press Make CV again to rebuild it."
            )
        if manual["skills"]:
            st.markdown("**Skills detected from your details**")
            st.markdown(
                "".join(f'<span class="pill">{s}</span>' for s in manual["skills"]),
                unsafe_allow_html=True,
            )
        with st.expander("Preview CV content — confirm every detail is included"):
            st.text(manual.get("text", ""))
        if st.button("⬇  Choose Download Format", type="primary", key="make_download"):
            download_format_dialog(manual)

    exit_button("exit_make")


if st.session_state.get("cv_ready") and screen in {"upload", "existing"}:
    meta = st.session_state["cv_meta"]
    st.divider()
    step(2, "Paste the job description")
    st.caption(f"Selected CV: {meta.get('original_name', 'CV.pdf')}")
    jd = st.text_area(
        "Job Description",
        height=260,
        label_visibility="collapsed",
        placeholder=(
            "Paste the company, role, responsibilities, required skills and "
            "experience here..."
        ),
    )

    if st.button("⚡  Generate ATS CV", type="primary"):
        if len(jd.strip()) < 50:
            st.error("Please paste the complete JD (at least 50 characters).")
        else:
            match = match_jd(st.session_state["cv_text"], jd)
            tailored = build_tailored_text(st.session_state["cv_text"], match)
            st.session_state["generated"] = {
                "match": match,
                "files": build_all_formats(tailored),
                "text": tailored,
            }

    generated = st.session_state.get("generated")
    if generated and "files" in generated:
        match = generated["match"]
        st.divider()
        step(3, "Your CV is ready")
        col1, col2 = st.columns(2)
        col1.metric("JD skill match", f"{match['score']}%")
        col2.metric("Verified skills", len(match["matched"]))

        if match["matched"]:
            st.markdown("**Matched from your CV**")
            st.markdown(
                "".join(f'<span class="pill">{s}</span>' for s in match["matched"]),
                unsafe_allow_html=True,
            )
        if match["missing"]:
            st.warning(
                "Required by the JD but not verified in your CV: "
                + ", ".join(match["missing"])
                + " — these skills were not added to the CV."
            )
        if not match["requested"]:
            st.info(
                "No known technical skills were identified in the JD. Your "
                "original CV was converted into an ATS-friendly one-column PDF."
            )

        with st.expander("Preview CV content"):
            st.text(generated.get("text", ""))

        if st.button("⬇  Choose Download Format", type="primary"):
            download_format_dialog(generated)

st.divider()
st.markdown(
    '<div class="privacy"><strong>Privacy:</strong> your CV and all generated '
    "files stay in the My_Agent folder on this computer. This app never logs in "
    "to a job portal or applies on your behalf.</div>",
    unsafe_allow_html=True,
)
