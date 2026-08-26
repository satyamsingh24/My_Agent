"""Standalone browser UI for creating truthful ATS CVs from a JD."""
from __future__ import annotations

import base64
from datetime import datetime
from pathlib import Path

import streamlit as st

from ai_providers import (
    PROVIDER_LABELS,
    apply_ai_response,
    provider_settings,
    run_ai_analysis,
    test_provider,
)
from cv_engine import (
    CV_TEMPLATES,
    ats_report,
    build_template_preview,
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
    restructure_to_reference,
    save_existing_cv,
    save_generated,
)

APP_PIN = "6932"
AI_SETTINGS_PASSWORD = "6932AI"
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
st.session_state.setdefault("ai_settings", provider_settings())


@st.cache_data(show_spinner=False)
def theme_css(theme_name: str) -> str:
    """Build the theme stylesheet once per theme instead of on every rerun."""
    t = THEMES[theme_name]
    return f"""
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
        .hero.lock-hero {{ padding-bottom: 3.4rem; }}
        .st-key-lock_pill {{
            margin-top: -3.35rem; margin-bottom: 1.2rem;
            position: relative; z-index: 2; text-align: center;
            animation: agentFadeUp .55s ease both;
        }}
        .st-key-lock_pill [data-testid="stElementContainer"] {{
            width: 100% !important; display: flex !important;
            justify-content: center !important;
        }}
        .st-key-lock_pill [data-testid="stButton"] {{
            display: flex; justify-content: center; width: auto;
        }}
        .st-key-lock_pill div.stButton > button {{
            width: auto !important; min-height: 0; padding: .5rem 1.15rem;
            border-radius: 999px; font-weight: 600; font-size: .9rem;
            background: rgba(255,255,255,.12); color: #eaf4ff !important;
            border: 1px solid rgba(255,255,255,.22); box-shadow: none;
        }}
        .st-key-lock_pill div.stButton > button:hover {{
            transform: translateY(-1px) scale(1.02);
            background: rgba(255,255,255,.22); color: #ffffff !important;
            border-color: rgba(255,255,255,.42); box-shadow: none;
        }}
        .st-key-lock_pill div.stButton > button p {{
            color: inherit !important; margin: 0;
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
        div.stButton > button, div.stDownloadButton > button,
        div[data-testid="stFormSubmitButton"] > button {{
            width: 100%; min-height: 3.1rem; border-radius: 14px;
            font-weight: 700; font-size: 1rem; letter-spacing: .01em;
            background: {t["panel"]}; color: {t["text"]};
            border: 1.5px solid {t["border"]};
            transition: transform .16s ease, box-shadow .16s ease,
                        background .16s ease, border-color .16s ease, color .16s ease;
        }}
        div.stButton > button:hover, div.stDownloadButton > button:hover,
        div[data-testid="stFormSubmitButton"] > button:hover {{
            transform: translateY(-2px);
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
            color: #ffffff !important; border-color: transparent;
            box-shadow: 0 8px 18px {t["accent"]}33;
        }}
        div.stButton > button:hover p, div.stDownloadButton > button:hover p,
        div[data-testid="stFormSubmitButton"] > button:hover p {{
            color: #ffffff !important;
        }}
        div.stButton > button:active, div.stDownloadButton > button:active,
        div[data-testid="stFormSubmitButton"] > button:active {{
            transform: translateY(-1px) scale(.99);
        }}
        div.stButton > button[kind="primary"],
        div.stDownloadButton > button[kind="primary"],
        div[data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"] {{
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
            color: #fff; border: none;
            box-shadow: 0 6px 16px {t["accent"]}2e;
        }}
        div.stButton > button[kind="primary"]:hover,
        div.stDownloadButton > button[kind="primary"]:hover,
        div[data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"]:hover {{
            background: linear-gradient(135deg, {t["accent2"]}, {t["accent"]});
            box-shadow: 0 10px 22px {t["accent2"]}3d;
        }}
        .option-card {{
            background: {t["panel"]}; border: 1.5px solid {t["border"]};
            border-radius: 18px; padding: 1.35rem 1.25rem 1.1rem;
            margin: .55rem 0 .3rem; box-shadow: {t["shadow"]};
            position: relative; overflow: hidden; transition: .2s ease;
            cursor: pointer;
        }}
        [class*="st-key-home_card_"] div.stButton > button {{
            display: flex; flex-direction: column; align-items: flex-start;
            justify-content: flex-start; gap: .1rem; text-align: left;
            white-space: normal; min-height: 13.5rem;
            padding: 1.35rem 1.25rem 1.1rem; border-radius: 18px;
            background: {t["panel"]}; color: {t["text"]};
            border: 1.5px solid {t["border"]}; box-shadow: {t["shadow"]};
            position: relative; overflow: hidden;
            animation: agentFadeUp .5s ease both;
        }}
        [class*="st-key-home_card_"] div.stButton > button::before {{
            content: ""; position: absolute; inset: 0 0 auto 0; height: 4px;
            background: linear-gradient(90deg, {t["accent"]}, {t["accent2"]});
            transform-origin: left center;
            animation: agentBarGrow .7s ease both;
        }}
        [class*="st-key-home_card_"] div.stButton > button:hover {{
            background: {t["panel"]} !important;
            border-color: {t["accent"]}; transform: translateY(-3px);
            box-shadow: 0 12px 24px {t["accent"]}2e;
        }}
        [class*="st-key-home_card_"] div.stButton > button p {{
            color: {t["text"]} !important; margin: 0;
        }}
        [class*="st-key-home_card_"] div.stButton > button:hover p {{
            color: {t["text"]} !important;
        }}
        [class*="st-key-home_card_"] div.stButton > button p:first-child:not(:last-child) {{
            width: 3rem; height: 3rem; border-radius: 14px; font-size: 1.4rem;
            display: flex; align-items: center; justify-content: center;
            background: linear-gradient(135deg, {t["accent"]}, {t["accent2"]});
            color: #fff !important; margin-bottom: .7rem;
            transition: transform .25s ease;
        }}
        [class*="st-key-home_card_"] div.stButton > button:hover p:first-child:not(:last-child) {{
            transform: translateY(-2px) scale(1.06);
            color: #fff !important;
        }}
        [class*="st-key-home_card_"] div.stButton > button p:nth-child(2) {{
            font-size: 1.12rem; font-weight: 750; margin-bottom: .3rem;
        }}
        [class*="st-key-home_card_"] div.stButton > button p:nth-child(3) {{
            font-size: .9rem; font-weight: 500; line-height: 1.5;
            color: {t["muted"]} !important;
        }}
        [class*="st-key-home_card_"] div.stButton > button:hover p:nth-child(3) {{
            color: {t["muted"]} !important;
        }}
        .st-key-home_card_existing div.stButton > button {{
            animation-delay: .09s;
        }}
        .st-key-home_card_make div.stButton > button {{
            animation-delay: .18s;
        }}
        .st-key-home_card_templates div.stButton > button {{
            animation-delay: .27s;
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
        .st-key-theme_toggle div.stButton > button,
        .st-key-ai_provider_button div.stButton > button {{
            width: 3.1rem; min-height: 3.1rem; padding: 0;
            border-radius: 50%; font-size: 1.35rem;
            background: {t["panel"]}; border: 1.5px solid {t["border"]};
            box-shadow: {t["shadow"]};
        }}
        .st-key-theme_toggle div.stButton > button:hover,
        .st-key-ai_provider_button div.stButton > button:hover {{
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
            display: flex;
            flex-direction: column;
            min-height: calc(100vh - 1.2rem);
        }}
        .st-key-side_help {{
            margin-top: auto;
            padding-bottom: .45rem;
        }}
        .side-desc {{
            margin: .55rem .1rem 0;
            font-size: .78rem;
            line-height: 1.5;
            color: {t["muted"]};
            text-align: left;
        }}
        .side-desc strong {{
            display: block;
            margin-bottom: .28rem;
            font-size: .8rem;
            letter-spacing: .06em;
            text-transform: uppercase;
            color: {t["text"]};
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

        .sheet-veil {{
            position: fixed; inset: 0; z-index: 900;
            background: rgba(6, 8, 18, .46);
            backdrop-filter: blur(3px);
            animation: sheetVeilIn .22s ease both;
        }}
        .st-key-pin_sheet {{
            position: fixed; left: 50%; bottom: 0; z-index: 901;
            width: min(400px, 94vw);
            padding: .8rem 1.15rem 1.25rem;
            background: {t["panel"]};
            border: 1px solid {t["border"]}; border-bottom: none;
            border-radius: 22px 22px 0 0;
            box-shadow: 0 -16px 42px rgba(6, 8, 18, .38);
            animation: sheetUp .3s cubic-bezier(.2, .74, .32, 1) both;
        }}
        .sheet-grip {{
            width: 44px; height: 5px; border-radius: 999px;
            background: {t["border"]}; margin: 0 auto .75rem;
        }}
        .sheet-head {{
            text-align: center; font-weight: 700; font-size: 1rem;
            margin-bottom: .7rem; color: {t["text"]};
        }}
        .st-key-pin_sheet div[data-testid="stTextInputRootElement"] {{
            background: {t["panel_soft"]}; border: 1.5px solid {t["border"]};
        }}
        .st-key-pin_sheet .stTextInput input {{
            text-align: center; letter-spacing: .45em; font-size: 1.15rem;
            min-height: 3.1rem;
        }}
        .st-key-pin_sheet .stTextInput input::placeholder {{
            letter-spacing: .3em; color: {t["muted"]};
        }}
        .st-key-pin_sheet div[data-testid="stForm"] {{ margin-bottom: .35rem; }}
        /* The hint would sit on top of the centred PIN dots. */
        .st-key-pin_sheet div[data-testid="InputInstructions"] {{ display: none; }}
        @keyframes sheetUp {{
            from {{ transform: translate(-50%, 100%); opacity: 0; }}
            to   {{ transform: translate(-50%, 0); opacity: 1; }}
        }}
        @keyframes sheetVeilIn {{
            from {{ opacity: 0; }}
            to   {{ opacity: 1; }}
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
        @keyframes agentBarGrow {{
            from {{ transform: scaleX(0); }}
            to   {{ transform: scaleX(1); }}
        }}

        .hero {{ animation: agentFadeUp .55s ease both; }}
        .step {{ animation: agentFadeUp .45s ease both; }}
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
        /* Streamlit paints the input shell from its own base theme, so the
           strip behind the reveal icon must follow the app theme instead. */
        div[data-testid="stTextInputRootElement"] {{
            background: {t["panel"]}; border: 1px solid {t["border"]};
            border-radius: 14px; overflow: hidden;
            transition: border-color .18s ease, box-shadow .18s ease;
        }}
        div[data-testid="stTextInputRootElement"]:focus-within {{
            border-color: {t["accent"]};
            box-shadow: 0 0 0 3px {t["accent"]}26;
        }}
        .stTextInput input {{
            background: transparent !important; border: none !important;
            color: {t["text"]};
        }}
        .stTextInput input::placeholder, .stTextArea textarea::placeholder {{
            color: {t["muted"]}; opacity: .85;
        }}
        .stTextInput button {{
            background: transparent !important; border: none !important;
            color: {t["muted"]} !important; min-height: 0 !important;
            width: auto !important;
        }}
        .stTextInput button:hover {{
            color: {t["accent"]} !important; transform: none !important;
            box-shadow: none !important;
        }}
        .stTextArea textarea {{
            transition: border-color .18s ease, box-shadow .18s ease;
        }}
        .stTextArea textarea:focus {{
            border-color: {t["accent"]} !important;
            box-shadow: 0 0 0 3px {t["accent"]}26;
        }}
        @media (prefers-reduced-motion: reduce) {{
            .hero, .step, .step .num, .card, .option-card,
            .option-card::before, .privacy, div[data-testid="stExpander"],
            div[data-testid="stMetric"], div[data-testid="stAlert"],
            .sheet-veil, .st-key-lock_pill {{
                animation: none;
            }}
            .st-key-pin_sheet {{
                animation: none; transform: translateX(-50%);
            }}
        }}
        </style>
        """


def paint(theme_name: str) -> None:
    st.markdown(theme_css(theme_name), unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def backdrop_photo_uri() -> str:
    """Optional user photo backdrop: drop a file in assets/background.<ext>."""
    for ext in ("jpg", "jpeg", "png", "webp"):
        path = ASSETS_DIR / f"background.{ext}"
        if path.exists():
            mime = "jpeg" if ext in {"jpg", "jpeg"} else ext
            encoded = base64.b64encode(path.read_bytes()).decode("ascii")
            return f"data:image/{mime};base64,{encoded}"
    return ""


@st.cache_data(show_spinner=False)
def _is_local_host(host: str) -> bool:
    name = host.split(":")[0].lower()
    return name in {"localhost", "127.0.0.1", "0.0.0.0", "::1", ""}


def is_hosted() -> bool:
    """True when the app is served to visitors instead of running locally."""
    try:
        host = st.context.headers.get("host", "")
    except Exception:
        return False
    return not _is_local_host(host)


def freeform_editor_url() -> str:
    """Static editor route for GitHub Pages/stlite and local Streamlit."""
    theme = st.session_state.get("theme", "Dark").lower()
    if is_hosted():
        return f"./static/editor/index.html?return=../../&theme={theme}"
    return f"/app/static/editor/index.html?return=../../../&theme={theme}"


def current_cv() -> tuple[dict, str, bytes] | None:
    """Hosted visitors keep their CV in their own session, never on disk."""
    if is_hosted():
        return st.session_state.get("session_cv")
    return load_existing_cv()


def store_cv(pdf_bytes: bytes, original_name: str, text: str) -> dict:
    if is_hosted():
        meta = {
            "original_name": original_name,
            "uploaded_at": datetime.now().isoformat(timespec="seconds"),
            "characters": len(text),
        }
        st.session_state["session_cv"] = (meta, text, pdf_bytes)
        return meta
    return save_existing_cv(pdf_bytes, original_name, text)


@st.cache_data(show_spinner=False)
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
            will-change: transform;
            animation: agentZoom 44s ease-in-out infinite alternate;
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
    """Front-page backdrop. Only transform/opacity animate, so the GPU does
    the work and clicks stay responsive."""
    t = THEMES[theme_name]
    blob = "26" if theme_name == "Dark" else "1c"
    grid = "0.05" if theme_name == "Dark" else "0.035"
    st.markdown(
        f"""
        <style>
        .stApp::before {{
            content: ""; position: fixed; inset: -25%;
            pointer-events: none; z-index: 0;
            background-image:
                radial-gradient(620px 420px at 18% 12%, {t["accent"]}{blob}, transparent 70%),
                radial-gradient(560px 400px at 82% 18%, {t["accent2"]}{blob}, transparent 70%),
                radial-gradient(520px 380px at 50% 92%, {t["accent"]}{blob}, transparent 70%),
                linear-gradient(rgba(125,140,220,{grid}) 1px, transparent 1px),
                linear-gradient(90deg, rgba(125,140,220,{grid}) 1px, transparent 1px);
            background-repeat: no-repeat, no-repeat, no-repeat, repeat, repeat;
            background-size: 90% 90%, 80% 80%, 95% 95%, 62px 62px, 62px 62px;
            will-change: transform;
            animation: agentDrift 40s ease-in-out infinite alternate;
        }}
        .block-container {{ position: relative; z-index: 1; }}
        .hero {{ position: relative; overflow: hidden; }}
        .hero::after {{
            content: ""; position: absolute; top: -60%; left: -35%;
            width: 45%; height: 220%; pointer-events: none;
            background: linear-gradient(
                100deg, transparent, rgba(255,255,255,.16), transparent);
            will-change: transform, opacity;
            animation: agentSheen 5s ease-in-out 2 both;
        }}
        @keyframes agentDrift {{
            0%   {{ transform: translate3d(0, 0, 0) scale(1); }}
            100% {{ transform: translate3d(2.5%, 2%, 0) scale(1.04); }}
        }}
        @keyframes agentSheen {{
            0%   {{ transform: translate3d(0, 0, 0) rotate(14deg); opacity: 0; }}
            20%  {{ opacity: .9; }}
            65%  {{ transform: translate3d(340%, 0, 0) rotate(14deg); opacity: 0; }}
            100% {{ transform: translate3d(340%, 0, 0) rotate(14deg); opacity: 0; }}
        }}
        @media (prefers-reduced-motion: reduce) {{
            .stApp::before, .hero::after {{ animation: none; }}
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
    for key in (
        "cv_meta",
        "cv_text",
        "cv_pdf",
        "cv_ready",
        "generated",
        "upload_generated",
        "upload_source",
        "existing_ats_source",
        "manual_ats_source",
    ):
        st.session_state.pop(key, None)
    st.session_state["change_cv"] = False


def privacy_note() -> str:
    if is_hosted():
        return (
            "<strong>Privacy:</strong> your CV stays in your own browser session "
            "and is never saved on the server or shared with other visitors. "
            "This app never logs in to a job portal or applies on your behalf."
        )
    return (
        "<strong>Privacy:</strong> your CV and all generated files stay in the "
        "My_Agent folder on this computer. This app never logs in to a job "
        "portal or applies on your behalf."
    )


def step(number: int, label: str) -> None:
    st.markdown(
        f'<div class="step"><span class="num">{number}</span>{label}</div>',
        unsafe_allow_html=True,
    )


def exit_button(key: str) -> None:
    if st.button("✕  Exit to Front Page", key=key):
        clear_cv_session()
        go("home")


@st.dialog("How to use My_AGENT")
def help_dialog() -> None:
    st.markdown(
        """
**Description**

Use this application to rebuild your CV in a more selective way for any
company’s job description (JD). It makes the CV ATS-friendly and ready
according to that JD, so when you apply, the company’s AI screening can
give your CV higher priority and your chance of being shortlisted increases.

My_AGENT also helps you write a short HR referral email. It does **not**
apply to jobs or log in to any job portal.

**Home page**
1. **Upload CV** — save a PDF CV so the app can use your details.
2. **Existing CV** — continue with the saved CV, paste a job description, then generate a matching ATS CV.
3. **Make your CV** — fill the form (photo is optional) and build a new ATS CV from only what you type.

**After you generate a CV**
Choose **PDF**, **DOCX**, or **XLSX**. Use PDF or DOCX when you apply. XLSX is only a reference copy.

**Left panel**
- **My Profile** — creator details and a reference CV download.
- **Cold Mail for Referral** — upload a CV and get a short email for HR (fresher / intern / experienced, based on that CV).
- **Settings** — log out and return to the PIN screen.

Read this once, then start from **Upload CV** or **Make your CV**.
        """
    )
    if st.button("Got it", type="primary", use_container_width=True):
        st.rerun()


def pin_sheet() -> None:
    """Small lock sheet that slides up from the bottom, like a phone lock."""
    st.markdown('<div class="sheet-veil"></div>', unsafe_allow_html=True)
    with st.container(key="pin_sheet"):
        st.markdown(
            '<div class="sheet-grip"></div>'
            '<div class="sheet-head">🔒 Enter your PIN</div>',
            unsafe_allow_html=True,
        )
        with st.form("pin_form", border=False):
            entered = st.text_input(
                "PIN",
                type="password",
                max_chars=8,
                key="pin_input",
                placeholder="••••",
                label_visibility="collapsed",
            )
            submitted = st.form_submit_button(
                "Unlock", type="primary", use_container_width=True
            )
        if submitted:
            if entered.strip() == APP_PIN:
                st.session_state["unlocked"] = True
                st.session_state["pin_sheet_open"] = False
                st.rerun()
            st.error("Incorrect PIN. Please try again.")
        if st.button("Cancel", use_container_width=True, key="pin_cancel"):
            st.session_state["pin_sheet_open"] = False
            st.rerun()


@st.dialog("Select AI Provider")
def ai_provider_dialog() -> None:
    """Password-protected, session-only provider configuration."""
    if not st.session_state.get("ai_settings_unlocked"):
        st.caption(
            "Enter the AI settings password. This is separate from the app PIN."
        )
        entered = st.text_input(
            "Password",
            type="password",
            key="ai_provider_password_input",
        )
        if st.button(
            "Unlock provider settings",
            type="primary",
            use_container_width=True,
            key="ai_provider_unlock",
        ):
            if entered == AI_SETTINGS_PASSWORD:
                st.session_state["ai_settings_unlocked"] = True
                st.rerun()
            else:
                st.error("Incorrect AI settings password.")
        if st.button(
            "Cancel",
            use_container_width=True,
            key="ai_provider_cancel_locked",
        ):
            st.session_state["show_ai_provider"] = False
            st.rerun()
        return

    settings = provider_settings(st.session_state.get("ai_settings"))
    providers = list(PROVIDER_LABELS)
    selected = str(settings.get("selected", "ollama"))
    selected_index = providers.index(selected) if selected in providers else 0

    st.warning(
        "The password protects this screen in the UI, but it is hardcoded in "
        "this public project and is not a security boundary."
    )
    with st.form("ai_provider_settings_form", clear_on_submit=False):
        provider_name = st.selectbox(
            "Select AI Provider",
            options=providers,
            index=selected_index,
            format_func=lambda name: PROVIDER_LABELS[name],
        )
        fallback = st.checkbox(
            "Automatically fall back to the other configured providers",
            value=bool(settings.get("fallback", True)),
        )
        st.caption(
            "Default chain starts with your selected provider. If it fails, "
            "the remaining configured providers are tried automatically."
        )

        st.markdown("**Ollama — free and local**")
        ollama_col, ollama_model_col = st.columns([1.25, 1])
        with ollama_col:
            ollama_url = st.text_input(
                "Ollama URL",
                value=str(settings.get("ollama_url", "")),
                placeholder="http://localhost:11434",
            )
        with ollama_model_col:
            ollama_model = st.text_input(
                "Ollama model",
                value=str(settings.get("ollama_model", "")),
                placeholder="llama3.2:3b",
            )

        st.markdown("**Gemini — free tier**")
        gemini_model = st.text_input(
            "Gemini model",
            value=str(settings.get("gemini_model", "")),
        )
        gemini_key = st.text_input(
            "Gemini API key",
            value=str(settings.get("gemini_key", "")),
            type="password",
            help="Kept only in this Streamlit session; never written to disk.",
        )

        st.markdown("**Groq — free tier**")
        groq_model = st.text_input(
            "Groq model",
            value=str(settings.get("groq_model", "")),
        )
        groq_key = st.text_input(
            "Groq API key",
            value=str(settings.get("groq_key", "")),
            type="password",
            help="Kept only in this Streamlit session; never written to disk.",
        )
        timeout = st.slider(
            "Provider timeout (seconds)",
            min_value=5,
            max_value=120,
            value=int(settings.get("timeout", 45)),
            step=5,
        )
        save_settings = st.form_submit_button(
            "Save for this session",
            type="primary",
            use_container_width=True,
        )

    if save_settings:
        st.session_state["ai_settings"] = provider_settings(
            {
                "selected": provider_name,
                "fallback": fallback,
                "ollama_url": ollama_url.strip(),
                "ollama_model": ollama_model.strip(),
                "gemini_model": gemini_model.strip(),
                "gemini_key": gemini_key.strip(),
                "groq_model": groq_model.strip(),
                "groq_key": groq_key.strip(),
                "timeout": timeout,
            }
        )
        st.session_state["ai_provider_status"] = (
            f"Saved: {PROVIDER_LABELS[provider_name]}"
        )
        st.rerun()

    status = st.session_state.get("ai_provider_status")
    if status:
        st.info(status)

    test_col, close_col = st.columns(2)
    with test_col:
        if st.button(
            "Test selected provider",
            use_container_width=True,
            key="ai_provider_test",
        ):
            current = provider_settings(st.session_state.get("ai_settings"))
            with st.spinner("Testing provider connection..."):
                ok, message = test_provider(
                    str(current["selected"]),
                    current,
                )
            st.session_state["ai_provider_status"] = (
                ("Connected: " if ok else "Connection failed: ") + message
            )
            st.rerun()
    with close_col:
        if st.button(
            "Close",
            use_container_width=True,
            key="ai_provider_close",
        ):
            st.session_state["show_ai_provider"] = False
            st.session_state["ai_settings_unlocked"] = False
            st.session_state.pop("ai_provider_password_input", None)
            st.rerun()

    st.caption(
        "Privacy: Ollama stays on this computer. Gemini or Groq receives the "
        "uploaded CV and JD when selected or used as a fallback. Cloud keys "
        "and settings are cleared when you log out."
    )


def align_cv_to_jd(
    cv_text: str,
    jd_text: str,
    extra_skills: list[str] | None = None,
    headline: str | None = None,
    ai_settings: dict | None = None,
) -> dict:
    """Apply the same deterministic + optional AI ATS pipeline to any CV."""
    original_match = match_jd(cv_text, jd_text)
    tailored = build_tailored_text(
        cv_text,
        original_match,
        extra_skills=extra_skills,
        headline=headline,
    )
    match = match_jd(tailored, jd_text)
    report = ats_report(tailored, jd_text, match)
    ai_result = run_ai_analysis(
        tailored,
        jd_text,
        report,
        allowed_skills=match["matched"],
        missing_skills=match["missing"],
        settings=ai_settings,
    )
    if ai_result["ok"] and ai_result["response"]:
        # Re-apply the reference layout so AI edits cannot disturb the
        # heading / bullet structure.
        tailored = restructure_to_reference(
            apply_ai_response(tailored, ai_result["response"])
        )
        match = match_jd(tailored, jd_text)
        report = ats_report(tailored, jd_text, match)

    return {
        "match": match,
        "report": report,
        "text": tailored,
        "added_skills": list(extra_skills or []),
        "ai": ai_result,
    }


def build_upload_docx(
    cv_text: str,
    jd_text: str,
    source_name: str,
    extra_skills: list[str] | None = None,
    headline: str | None = None,
    ai_settings: dict | None = None,
) -> dict:
    """Rebuild an uploaded CV as an ATS DOCX aligned to the pasted JD."""
    result = align_cv_to_jd(
        cv_text,
        jd_text,
        extra_skills=extra_skills,
        headline=headline,
        ai_settings=ai_settings,
    )
    stem = random_output_stem()
    docx_bytes = build_ats_docx(
        result["text"], title=stem, template="reference"
    )
    filename = stem + ".docx"
    if not is_hosted():
        save_generated(docx_bytes, filename)
    return {
        **result,
        "data": docx_bytes,
        "filename": filename,
        "source_name": source_name,
    }


def build_all_formats(
    text: str,
    stem: str | None = None,
    photo_bytes: bytes | None = None,
    template: str = "reference",
) -> dict:
    stem = stem or random_output_stem()
    pdf_bytes = build_ats_pdf(
        text, title=stem, photo_bytes=photo_bytes, template=template
    )
    docx_bytes = build_ats_docx(
        text, title=stem, photo_bytes=photo_bytes, template=template
    )
    xlsx_bytes = build_cv_xlsx(
        text, title=stem, photo_bytes=photo_bytes, template=template
    )
    hosted = is_hosted()

    def keep(data: bytes, filename: str) -> str:
        if hosted:
            return ""
        return str(save_generated(data, filename))

    return {
        "pdf": {
            "label": "PDF",
            "data": pdf_bytes,
            "filename": stem + ".pdf",
            "mime": "application/pdf",
            "path": keep(pdf_bytes, stem + ".pdf"),
        },
        "docx": {
            "label": "DOCX",
            "data": docx_bytes,
            "filename": stem + ".docx",
            "mime": (
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            "path": keep(docx_bytes, stem + ".docx"),
        },
        "xlsx": {
            "label": "XLSX",
            "data": xlsx_bytes,
            "filename": stem + ".xlsx",
            "mime": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "path": keep(xlsx_bytes, stem + ".xlsx"),
        },
    }


def build_manual_jd_result(
    base_text: str,
    jd_text: str,
    photo_bytes: bytes | None,
    template: str,
    signature: str,
    template_label: str,
    extra_skills: list[str] | None = None,
    headline: str | None = None,
) -> dict:
    """Apply the Option 1 ATS pipeline while retaining all manual formats."""
    result = align_cv_to_jd(
        base_text,
        jd_text,
        extra_skills=extra_skills,
        headline=headline,
        ai_settings=st.session_state.get("ai_settings"),
    )
    return {
        **result,
        "files": build_all_formats(
            result["text"],
            photo_bytes=photo_bytes,
            template=template,
        ),
        "skills": result["match"]["matched"],
        "signature": signature,
        "template": template,
        "template_label": template_label,
    }


def render_ats_overview(result: dict) -> None:
    """Shared ATS/AI report used by Upload, Existing and Make CV flows."""
    match = result["match"]
    report = result["report"]
    title = report["title"]

    if match["missing"]:
        st.warning(
            "Missing from your CV but requested in the JD: "
            + ", ".join(match["missing"])
        )
        st.caption(
            "Nothing missing is added unless you truthfully confirm it."
        )

    overall_col, skills_col, title_col, format_col = st.columns(4)
    overall_col.metric("Overall ATS score", f"{report['overall']}%")
    skills_col.metric("Skill match", f"{report['skills_score']}%")
    title_col.metric(
        "Title match",
        "N/A" if title["score"] is None else f"{title['score']}%",
    )
    format_col.metric("Format", f"{report['format_score']}%")

    if report["required_years"]:
        st.caption(
            f"Experience: JD asks for about {report['required_years']} years; "
            f"CV dates add up to {report['cv_years']} years."
        )
    else:
        st.caption(
            f"Experience: JD states no year requirement; CV dates add up to "
            f"{report['cv_years']} years."
        )
    if title["jd_title"]:
        st.caption(
            f"JD title: {title['jd_title']} · CV headline: "
            f"{title['cv_title'] or 'not detected'}"
        )

    if match["matched"]:
        st.markdown("**Verified JD skills prioritised in the CV**")
        st.markdown(
            "".join(
                f'<span class="pill">{skill}</span>'
                for skill in match["matched"]
            ),
            unsafe_allow_html=True,
        )
    if result.get("added_skills"):
        st.info(
            "Added because you confirmed them: "
            + ", ".join(result["added_skills"])
        )

    ai_result = result.get("ai", {})
    if ai_result.get("ok"):
        provider_name = str(ai_result.get("provider", ""))
        st.success(
            "AI suggestions prepared by "
            + PROVIDER_LABELS.get(provider_name, provider_name.title())
            + "."
        )
        ai_response = ai_result.get("response") or {}
        if ai_response.get("suggestions"):
            st.markdown("**AI suggestions**")
            for suggestion in ai_response["suggestions"]:
                st.markdown(f"- {suggestion}")
        for warning in ai_response.get("validation_warnings", []):
            st.warning(warning)
    else:
        st.warning(
            "No AI provider was available. Deterministic ATS rules still "
            "generated the CV safely."
        )

    if report["suggestions"]:
        st.markdown("**Improve your chances**")
        for tip in report["suggestions"]:
            st.markdown(f"- {tip}")


@st.cache_data(show_spinner=False)
def template_preview(text: str, template: str, width: int = 430) -> bytes:
    return build_template_preview(text, template, width=width)


def full_width_image(image: bytes) -> None:
    """Streamlit renamed the image sizing argument in 1.49, and the browser
    build still ships an older release, so pick whichever it understands."""
    try:
        version = tuple(int(part) for part in st.__version__.split(".")[:2])
    except ValueError:
        version = (0, 0)
    if version >= (1, 49):
        st.image(image, width="stretch")
    else:
        st.image(image, use_container_width=True)


SAMPLE_CV_SECTIONS = {
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


@st.cache_data(show_spinner=False)
def sample_cv_text() -> str:
    """Neutral example CV used only to show what each template looks like."""
    return compose_cv_text(
        name="Your Name",
        email="you@example.com",
        phone="98765 43210",
        city="Indore",
        state="Madhya Pradesh",
        headline="Python Developer",
        sections=SAMPLE_CV_SECTIONS,
    )


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
        <div class="hero lock-hero">
          <h1>Welcome to My_AGENT</h1>
          <div class="by">Created by Satyam Singh Bhadoriya</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.container(key="lock_pill"):
        if st.button(
            "🔒 Locked — tap unlock and enter your PIN",
            key="open_pin_sheet",
        ):
            st.session_state["pin_sheet_open"] = True
            st.rerun()
    if st.session_state.get("pin_sheet_open"):
        pin_sheet()
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
    with st.container(key="side_help"):
        if st.button("❓  How to use", use_container_width=True, key="side_howto"):
            help_dialog()
        st.markdown(
            """
            <div class="side-desc">
              <strong>Description</strong>
              Use this application to rebuild your CV in a more selective way
              for any company’s job description (JD). The CV is made
              ATS-friendly and aligned with that JD, so when you apply, the
              company’s AI screening can give your CV higher priority and your
              chance of being shortlisted increases.
            </div>
            """,
            unsafe_allow_html=True,
        )

is_dark = st.session_state["theme"] == "Dark"
spacer, ai_col, toggle_col = st.columns([7, 1, 1])
with ai_col:
    with st.container(key="ai_provider_button"):
        if st.button(
            "AI",
            help="Select and configure the free AI provider",
            key="open_ai_provider",
        ):
            st.session_state["show_ai_provider"] = True
with toggle_col:
    with st.container(key="theme_toggle"):
        if st.button(
            "☀" if is_dark else "🌙",
            help="Switch to Light mode" if is_dark else "Switch to Dark mode",
        ):
            st.session_state["theme"] = "Light" if is_dark else "Dark"
            st.rerun()

if st.session_state.get("show_ai_provider"):
    ai_provider_dialog()

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

existing = current_cv()

if screen == "home":
    step(1, "Choose an option")
    upload_note = (
        "Upload a CV PDF. It stays in your own browser session only."
        if is_hosted()
        else "Upload a new CV PDF. Your details are saved on this computer."
    )
    left, right, third, fourth = st.columns(4)
    with left:
        with st.container(key="home_card_upload"):
            if st.button(
                f"⬆\n\n1. Upload CV\n\n{upload_note}",
                key="home_upload",
                use_container_width=True,
            ):
                clear_cv_session()
                go("upload")
    with right:
        saved_name = (
            existing[0].get("original_name", "CV.pdf")
            if existing
            else "No CV saved yet"
        )
        with st.container(key="home_card_existing"):
            if st.button(
                f"📁\n\n2. Existing CV\n\nSaved: {saved_name}",
                key="home_existing",
                use_container_width=True,
            ):
                clear_cv_session()
                go("existing")
    with third:
        with st.container(key="home_card_make"):
            if st.button(
                "✍\n\n3. Make your CV\n\n"
                "Fill in your details and build a simple ATS CV.",
                key="home_make",
                use_container_width=True,
            ):
                clear_cv_session()
                st.session_state.pop("manual_generated", None)
                st.session_state.pop("make_template", None)
                go("make")
    with fourth:
        with st.container(key="home_card_templates"):
            if st.button(
                "🎨\n\n4. Templates\n\n"
                "Choose a template or open the freeform CV editor.",
                key="home_templates",
                use_container_width=True,
            ):
                clear_cv_session()
                st.session_state.pop("manual_generated", None)
                go("templates_hub")

elif screen == "templates_hub":
    step(1, "Templates and CV Editor")
    st.caption(
        "Build a structured ATS-friendly CV from a template, or freely edit "
        "your own selectable-text PDF in the visual editor."
    )
    template_choice, editor_choice = st.columns(2)
    with template_choice:
        with st.container(key="option4_template_choice"):
            st.markdown("### 1. Existing Templates")
            st.write(
                "Choose from 10 ATS-friendly designs, fill in your details and "
                "target JD, then download PDF, DOCX or XLSX."
            )
            if st.button(
                "View 10 templates",
                key="open_existing_templates",
                type="primary",
                use_container_width=True,
            ):
                go("templates")
    with editor_choice:
        with st.container(key="option4_editor_choice"):
            st.markdown("### 2. Edit Your CV")
            st.write(
                "Upload your own PDF, edit text and design elements, switch "
                "templates, and download the result as PDF or Word."
            )
            st.link_button(
                "Open freeform CV editor",
                freeform_editor_url(),
                use_container_width=True,
            )
            st.caption(
                "Your CV stays in this browser. Freeform columns and graphics "
                "may be less reliable in company ATS systems."
            )
    exit_button("exit_templates_hub")

elif screen == "templates":
    step(1, "Choose a template")
    st.caption(
        "Choose from 10 free ATS-friendly designs. Each preview uses example "
        "content; after selection, fill in your own details and target JD."
    )
    sample = sample_cv_text()
    # Drawing ten thumbnails is far slower in the browser build, so the hosted
    # app renders them smaller.
    preview_width = 260 if is_hosted() else 430
    template_items = list(CV_TEMPLATES.items())
    for row_start in range(0, len(template_items), 3):
        row_items = template_items[row_start : row_start + 3]
        columns = st.columns(3)
        for (template, design), column in zip(row_items, columns):
            with column:
                full_width_image(
                    template_preview(sample, template, preview_width)
                )
                st.markdown(f"**{design['label']}**")
                st.caption(design["description"])
                if st.button(
                    f"Use {design['label']}",
                    type="primary" if template == "reference" else "secondary",
                    use_container_width=True,
                    key=f"pick_template_{template}",
                ):
                    st.session_state["make_template"] = template
                    st.session_state.pop("manual_generated", None)
                    go("make")
    exit_button("exit_templates")

elif screen == "upload":
    step(1, "Upload CV and add the job description")
    st.caption(
        "Upload a selectable-text PDF. My_AGENT will rebuild its content as a "
        "clean ATS-friendly DOCX and prioritise only skills already verified "
        "in your CV."
    )
    with st.form("upload_cv_jd_form", clear_on_submit=False):
        uploaded = st.file_uploader(
            "Your CV (PDF)",
            type=["pdf"],
            help=(
                "The PDF must contain selectable text. Scanned image PDFs are "
                "not supported."
            ),
        )
        upload_jd = st.text_area(
            "Job Description",
            height=260,
            placeholder=(
                "Paste the company, role, responsibilities, required skills "
                "and experience here..."
            ),
        )
        upload_submit = st.form_submit_button(
            "⚡ Convert to ATS-friendly DOCX",
            type="primary",
            use_container_width=True,
        )

    if upload_submit:
        if uploaded is None:
            st.error("Please upload your CV in PDF format.")
        elif len(upload_jd.strip()) < 50:
            st.error("Please paste the complete JD (at least 50 characters).")
        else:
            try:
                file_bytes = uploaded.getvalue()
                with st.spinner("Reading the CV and preparing your ATS DOCX..."):
                    extracted = extract_pdf_text(file_bytes)
                    meta = store_cv(file_bytes, uploaded.name, extracted)
                    activate_cv(meta, extracted, file_bytes)
                    st.session_state["upload_source"] = {
                        "cv_text": extracted,
                        "jd_text": upload_jd,
                        "name": uploaded.name,
                    }
                    st.session_state["upload_generated"] = build_upload_docx(
                        extracted,
                        upload_jd,
                        uploaded.name,
                        ai_settings=st.session_state.get("ai_settings"),
                    )
            except Exception as exc:
                st.error(str(exc))

    upload_source = st.session_state.get("upload_source")
    upload_result = st.session_state.get("upload_generated")
    if upload_source and upload_result:
        match = upload_result["match"]
        report = upload_result["report"]
        title = report["title"]
        st.divider()
        step(2, "How an ATS will read this CV")
        st.success(
            f"{upload_result['source_name']} was converted and aligned to the JD."
        )

        if match["missing"]:
            st.warning(
                "Missing from your CV but requested in the JD: "
                + ", ".join(match["missing"])
            )
            st.markdown(
                "".join(
                    f'<span class="pill">{skill}</span>'
                    for skill in match["missing"]
                ),
                unsafe_allow_html=True,
            )
            st.caption(
                "These are suggestions only. Nothing missing is added unless "
                "you confirm it below."
            )

        overall_col, skills_col, title_col, format_col = st.columns(4)
        overall_col.metric("Overall ATS score", f"{report['overall']}%")
        skills_col.metric("Skill match", f"{report['skills_score']}%")
        title_col.metric(
            "Title match",
            "N/A" if title["score"] is None else f"{title['score']}%",
        )
        format_col.metric("Format", f"{report['format_score']}%")

        if report["required_years"]:
            st.caption(
                f"Experience: the JD asks for about {report['required_years']} "
                f"years and your CV dates add up to {report['cv_years']} years."
            )
        else:
            st.caption(
                f"Experience: the JD states no year requirement. Your CV dates "
                f"add up to {report['cv_years']} years."
            )
        if title["jd_title"]:
            st.caption(
                f"JD title: {title['jd_title']} · Your headline: "
                f"{title['cv_title'] or 'not detected'}"
            )

        if match["matched"]:
            st.markdown("**Verified skills prioritised in the DOCX**")
            st.markdown(
                "".join(
                    f'<span class="pill">{skill}</span>'
                    for skill in match["matched"]
                ),
                unsafe_allow_html=True,
            )
        if upload_result["added_skills"]:
            st.info(
                "Added because you confirmed them: "
                + ", ".join(upload_result["added_skills"])
            )
        if not match["requested"]:
            st.info(
                "No skill keywords were identified in the JD. The CV was still "
                "converted into the ATS-friendly reference format."
            )

        ai_result = upload_result.get("ai", {})
        if ai_result.get("ok"):
            provider_name = str(ai_result.get("provider", ""))
            st.success(
                "AI suggestions prepared by "
                + PROVIDER_LABELS.get(provider_name, provider_name.title())
                + "."
            )
            ai_response = ai_result.get("response") or {}
            if ai_response.get("suggestions"):
                st.markdown("**AI suggestions**")
                for suggestion in ai_response["suggestions"]:
                    st.markdown(f"- {suggestion}")
            for warning in ai_response.get("validation_warnings", []):
                st.warning(warning)
        else:
            st.warning(
                "No AI provider was available. The ATS rules still generated "
                "your DOCX safely; configure a provider from the AI button."
            )

        attempts = ai_result.get("attempts", [])
        if attempts:
            with st.expander("AI provider attempts"):
                for attempt in attempts:
                    label = PROVIDER_LABELS.get(
                        attempt["provider"], attempt["provider"].title()
                    )
                    detail = (
                        f" — {attempt['detail']}" if attempt.get("detail") else ""
                    )
                    st.write(f"{label}: {attempt['status']}{detail}")

        if report["suggestions"]:
            st.markdown("**Improve your chances**")
            for tip in report["suggestions"]:
                st.markdown(f"- {tip}")

        if match["missing"] or title["jd_title"]:
            st.divider()
            step(3, "Fix the gaps, then rebuild")
        confirmed_skills: list[str] = []
        if match["missing"]:
            confirmed_skills = st.multiselect(
                "Skills the JD wants that are not in your CV",
                options=match["missing"],
                key="upload_confirm_skills",
                help=(
                    "Tick only what you genuinely have. Anything you tick is "
                    "added to the CV; the rest is never invented for you."
                ),
            )
            st.caption(
                "Never tick a skill you cannot defend in an interview — a "
                "recruiter will verify it."
            )
        use_jd_title = False
        if title["jd_title"] and (title["score"] or 0) < 100:
            use_jd_title = st.checkbox(
                f"Set my CV headline to \"{title['jd_title']}\"",
                key="upload_use_jd_title",
                help=(
                    "Use this only when your actual role matches the JD title."
                ),
            )
        if match["missing"] or title["jd_title"]:
            if st.button(
                "🔁 Rebuild DOCX with my confirmations",
                key="upload_rebuild",
                use_container_width=True,
            ):
                try:
                    with st.spinner("Rebuilding your ATS DOCX..."):
                        st.session_state["upload_generated"] = build_upload_docx(
                            upload_source["cv_text"],
                            upload_source["jd_text"],
                            upload_source["name"],
                            extra_skills=confirmed_skills,
                            headline=title["jd_title"] if use_jd_title else None,
                            ai_settings=st.session_state.get("ai_settings"),
                        )
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        st.divider()
        step(4, "Download")
        with st.expander("Preview ATS CV content"):
            st.text(upload_result["text"])
        st.download_button(
            "⬇ Download ATS DOCX",
            data=upload_result["data"],
            file_name=upload_result["filename"],
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            type="primary",
            use_container_width=True,
            key="upload_docx_download",
        )
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
                        new_meta = store_cv(
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
    chosen_template = st.session_state.get("make_template", "reference")
    if chosen_template not in CV_TEMPLATES:
        chosen_template = "reference"
        st.session_state["make_template"] = chosen_template
    chosen_design = CV_TEMPLATES[chosen_template]
    step(1, "Make your CV")
    st.caption(
        "CV detail fields are optional, but a target JD is required for ATS "
        "tailoring. The more details you add, the more complete your CV will be."
    )
    if st.session_state.get("make_template"):
        badge, change = st.columns([3, 1])
        badge.success(f"Selected template: {chosen_design['label']}")
        if change.button("Change template", key="make_change_template"):
            st.session_state.pop("manual_generated", None)
            go("templates")
    else:
        st.info(
            "This builds a simple CV in the reference design. For other designs, "
            "use **Templates** on the front page."
        )
    # A form submits every field together, so one click never loses text that
    # was still being typed.
    with st.form("make_cv_form", clear_on_submit=False, border=False):
        form_left, form_right = st.columns(2)
        with form_left:
            full_name = st.text_input(
                "Name", key="make_name", placeholder="Your full name"
            )
            phone = st.text_input(
                "Number", key="make_phone", placeholder="Mobile number"
            )
        with form_right:
            email = st.text_input(
                "Email", key="make_email", placeholder="you@example.com"
            )
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
            placeholder=(
                "DevOps Engineer, Java Full Stack Developer, Python Developer..."
            ),
        )
        profile_photo = st.file_uploader(
            "Profile photo (optional)",
            type=["jpg", "jpeg", "png", "webp"],
            key="make_photo",
            help="If added, the photo is placed at the top-right of the CV.",
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
                placeholder=(
                    "Role\nCompany, city\nJoining date - end date\n"
                    "- What you built or handled"
                ),
            )
        with proj_col:
            projects = st.text_area(
                "Projects",
                key="make_projects",
                height=140,
                placeholder="Project name, what you built, stack",
            )
        edu_col, qual_col, cert_col = st.columns(3)
        with edu_col:
            education = st.text_area(
                "Education",
                key="make_education",
                height=110,
                placeholder="Degree, college, year",
            )
        with qual_col:
            qualification = st.text_area(
                "Qualification",
                key="make_qualification",
                height=110,
                placeholder="12th / 10th board, school, percentage, year",
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
                "Extra notes, additional skills, college or home address — "
                "anything missed in the fields above"
            ),
        )
        st.caption(
            "Note: this box accepts extra skills, experience, college details, "
            "or your home address. Nothing here is mandatory."
        )
        st.markdown("**Target Job Description — required for ATS tailoring**")
        make_jd = st.text_area(
            "Target Job Description",
            key="make_jd",
            height=220,
            label_visibility="collapsed",
            placeholder=(
                "Paste the company, role, responsibilities, required skills "
                "and experience here..."
            ),
        )
        st.caption(
            "The same ATS rules used by Upload CV will compare this JD with "
            "your form details."
        )
        make_clicked = st.form_submit_button(
            "⚡  Make CV", type="primary", use_container_width=True
        )
    photo_bytes = profile_photo.getvalue() if profile_photo is not None else None
    if photo_bytes:
        st.caption(
            "The photo appears at the top-right of the PDF and Word CV. "
            "Some ATS software ignore images, so the rest stays text-based."
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
        "qualification": qualification,
        "certifications": certifications,
        "internships": internships,
        "achievements": achievements,
        "languages": languages,
        "details": details,
        "job_description": make_jd,
        "photo": (
            f"{profile_photo.name}:{len(photo_bytes)}" if photo_bytes else ""
        ),
    }
    signature = repr(form_values)

    if make_clicked:
        if len(make_jd.strip()) < 50:
            st.error(
                "Please paste the complete target JD (at least 50 characters)."
            )
        else:
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
                        "qualification": qualification,
                        "certifications": certifications,
                        "internships": internships,
                        "achievements": achievements,
                        "languages": languages,
                    },
                )
                with st.spinner(
                    "Analysing the JD and preparing PDF, DOCX, and XLSX..."
                ):
                    st.session_state["manual_ats_source"] = {
                        "base_text": composed,
                        "jd_text": make_jd,
                        "photo_bytes": photo_bytes,
                        "template": chosen_template,
                        "signature": signature,
                        "template_label": chosen_design["label"],
                    }
                    st.session_state["manual_generated"] = (
                        build_manual_jd_result(
                            composed,
                            make_jd,
                            photo_bytes,
                            chosen_template,
                            signature,
                            chosen_design["label"],
                        )
                    )
            except Exception as exc:
                st.error(str(exc))

    manual = st.session_state.get("manual_generated")
    manual_source = st.session_state.get("manual_ats_source")
    if manual and manual_source:
        st.divider()
        step(2, "How an ATS will read this CV")
        if manual.get("template_label"):
            st.success(f"Template: {manual['template_label']}")
        if manual.get("signature") != signature:
            st.warning(
                "Your form details have changed. The CV below uses the older "
                "details — press Make CV again to rebuild it."
            )
        render_ats_overview(manual)

        match = manual["match"]
        title = manual["report"]["title"]
        confirmed_skills: list[str] = []
        if match["missing"] or title["jd_title"]:
            st.divider()
            step(3, "Fix the gaps, then rebuild")
        if match["missing"]:
            confirmed_skills = st.multiselect(
                "Skills the JD wants that are not in your details",
                options=match["missing"],
                key="make_confirm_skills",
                help="Tick only skills you genuinely have.",
            )
        use_jd_title = False
        if title["jd_title"] and (title["score"] or 0) < 100:
            use_jd_title = st.checkbox(
                f"Set my CV headline to \"{title['jd_title']}\"",
                key="make_use_jd_title",
                help="Use only when this truthfully describes your role.",
            )
        if match["missing"] or title["jd_title"]:
            if st.button(
                "🔁 Rebuild CV with my confirmations",
                key="make_ats_rebuild",
                use_container_width=True,
            ):
                try:
                    with st.spinner(
                        "Rebuilding ATS PDF, DOCX, and XLSX..."
                    ):
                        st.session_state["manual_generated"] = (
                            build_manual_jd_result(
                                manual_source["base_text"],
                                manual_source["jd_text"],
                                manual_source["photo_bytes"],
                                manual_source["template"],
                                manual_source["signature"],
                                manual_source["template_label"],
                                extra_skills=confirmed_skills,
                                headline=(
                                    title["jd_title"]
                                    if use_jd_title
                                    else None
                                ),
                            )
                        )
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        st.divider()
        with st.expander("Preview CV content — confirm every detail is included"):
            st.text(manual.get("text", ""))
        if st.button("⬇  Choose Download Format", type="primary", key="make_download"):
            download_format_dialog(manual)

    exit_button("exit_make")


if st.session_state.get("cv_ready") and screen == "existing":
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
            try:
                with st.spinner("Analysing CV, JD and AI suggestions..."):
                    st.session_state["existing_ats_source"] = {
                        "cv_text": st.session_state["cv_text"],
                        "jd_text": jd,
                        "name": meta.get("original_name", "Existing_CV.pdf"),
                    }
                    st.session_state["generated"] = build_upload_docx(
                        st.session_state["cv_text"],
                        jd,
                        meta.get("original_name", "Existing_CV.pdf"),
                        ai_settings=st.session_state.get("ai_settings"),
                    )
            except Exception as exc:
                st.error(str(exc))

    generated = st.session_state.get("generated")
    existing_source = st.session_state.get("existing_ats_source")
    if generated and existing_source and "data" in generated:
        st.divider()
        step(3, "How an ATS will read this CV")
        render_ats_overview(generated)

        match = generated["match"]
        title = generated["report"]["title"]
        confirmed_skills: list[str] = []
        if match["missing"] or title["jd_title"]:
            st.divider()
            step(4, "Fix the gaps, then rebuild")
        if match["missing"]:
            confirmed_skills = st.multiselect(
                "Skills the JD wants that are not in your saved CV",
                options=match["missing"],
                key="existing_confirm_skills",
                help="Tick only skills you genuinely have.",
            )
        use_jd_title = False
        if title["jd_title"] and (title["score"] or 0) < 100:
            use_jd_title = st.checkbox(
                f"Set my CV headline to \"{title['jd_title']}\"",
                key="existing_use_jd_title",
                help="Use only when your actual role matches this title.",
            )
        if match["missing"] or title["jd_title"]:
            if st.button(
                "🔁 Rebuild saved CV with my confirmations",
                key="existing_rebuild",
                use_container_width=True,
            ):
                try:
                    with st.spinner("Rebuilding your ATS DOCX..."):
                        st.session_state["generated"] = build_upload_docx(
                            existing_source["cv_text"],
                            existing_source["jd_text"],
                            existing_source["name"],
                            extra_skills=confirmed_skills,
                            headline=(
                                title["jd_title"] if use_jd_title else None
                            ),
                            ai_settings=st.session_state.get("ai_settings"),
                        )
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        st.divider()
        with st.expander("Preview ATS CV content"):
            st.text(generated.get("text", ""))
        st.download_button(
            "⬇ Download ATS DOCX",
            data=generated["data"],
            file_name=generated["filename"],
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            type="primary",
            use_container_width=True,
            key="existing_docx_download",
        )

st.divider()
st.markdown(
    f'<div class="privacy">{privacy_note()}</div>',
    unsafe_allow_html=True,
)
