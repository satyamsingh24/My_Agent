# My_AGENT — ATS CV Builder

A standalone Streamlit app for building ATS-friendly CVs. Python is enough to run it locally; no extra desktop tools are required.

## Access

The app opens on a lock screen. Tap **Locked — tap unlock and enter your PIN**, then enter the PIN in the sheet that slides up from the bottom. The PIN is set in `app.py` and is not listed here.

## Features

0. **Dark / Light mode** toggle in the top-right, and an **Exit to Front Page** button on every screen.
1. **Upload CV** — upload any selectable-text PDF (DevOps, Python, Java, and similar roles).
2. CV text is stored in local files on a local run.
3. **Existing CV** — continue or replace the saved CV, then use the full job-description, ATS score, AI/fallback, and truthful rebuild workflow.
4. Paste a complete job description.
5. The app detects skills that appear in both the CV and the job description.
6. Attractive ATS-safe one-column layout: readable Arial/Helvetica fonts, blue section headings, clean spacing, no columns, icons, or tables.
7. Choose **PDF**, **DOCX**, or **XLSX** from the download popup.
8. **Make your CV** — build a new CV from a form without an existing file, paste a target job description, and use the same ATS score, AI/fallback, and confirmation workflow. Output is available as PDF, DOCX, and XLSX.
9. The reference resume is used only for **layout and heading order**. Content comes only from the details you provide, not from another person's CV.
10. Even with few details, the CV still reads professionally: the role and skills produce a summary, grouped SKILLS lines (Cloud / DevOps / Databases, and similar), and CORE COMPETENCIES bullets built from your own keywords.
11. Changing form details shows a warning — press **Make CV** again, otherwise the previous CV stays on screen.
12. The front page uses a professional animated background (moving gradient glow, a slow grid, and a hero sheen) in pure CSS, with no internet or extra files. To use your own image, put `assets/background.jpg` (or `.png` / `.webp`) in the project; the app dims it and slow-zooms it as a backdrop. Remove the file to return to the CSS animation.
13. The sidebar **My Profile** page ships the owner's passport-size photo (`assets/owner_photo.jpg`) and basic details in the code, and **My CV for Reference** downloads the reference resume.
14. **Cold Mail for Referral** in the left sidebar reads a CV and writes a short HR email. Fresher, internship-only, and experienced profiles get different truthful wording. Skills or experience that are not in the CV are not added.
15. **Settings** in the sidebar logs out, locks the session, and returns to the PIN screen.
16. The top-right **AI** button opens provider settings behind a separate password (also set in `app.py`, not listed here). Ollama is local-first. If it fails, configured Gemini and Groq free-tier providers can be tried. Cloud API keys stay only in current session memory and are cleared on logout.
17. **Templates → Edit Your CV** opens a browser-only freeform editor. Import a selectable-text PDF, then move, resize, rotate, and edit text boxes. Use fonts, colours, shapes, lines, and images, apply any of the 10 template previews, undo/redo, and export as PDF or DOCX.

Every generated CV follows the layout of `profile/Satyam_Dev_Resume_ATS.pdf`: Name → Role → Contact → SUMMARY → SKILLS → PROFESSIONAL EXPERIENCE → INTERNSHIPS → PROJECTS → EDUCATION → CERTIFICATIONS. Declaration and References blocks are dropped because they do not help ATS parsers.

PDF and DOCX are ATS-friendly. XLSX is only a reference or editing copy; do not upload an Excel CV to a job portal.

Skill matching covers DevOps and cloud as well as Java full stack, Python, frontend, backend, mobile, data/AI, testing, and security technologies.

The app does not invent skills or experience. A requirement that appears in the job description but is not verified in the uploaded CV is listed as missing.

## Free AI providers

The Upload CV flow runs deterministic ATS checks first, then optional AI wording suggestions. The default provider chain is:

1. **Ollama** — free and local; the CV does not leave this computer.
2. **Gemini** — free tier; enter the API key in the AI settings dialog.
3. **Groq** — free tier; enter the API key in the AI settings dialog.

For Ollama, install [Ollama](https://ollama.com/), pull a model, and start the server:

```powershell
ollama pull llama3.2:3b
ollama serve
```

The AI settings password is hardcoded in the public source, so it is not a real security boundary. If Gemini or Groq is selected or used as a fallback, the uploaded CV and job description are sent to that cloud provider. The AI does not add missing job-description skills on its own; you must tick them in the truthful confirmation list first.

## Freeform CV editor

Open **4. Templates** on the home page:

1. **Existing Templates** starts the structured, job-description-aligned, ATS-safe CV flow.
2. **Edit Your CV** opens the visual editor. Use **Import PDF** for a selectable-text CV, then use the labelled toolbar for text, insert, arrange, history, zoom, and page tools, and choose a format under **Export**.

The left panel shows a real page preview for each template. Clicking a preview rebuilds the imported CV as editable objects in that design. Switching templates changes only layout and style; the app does not invent skills or experience. Preview images live in `static/editor/previews/` and can be regenerated with `python -m scripts.export_template_previews` after `CV_TEMPLATES` changes.

The editor header has its own dark/light switch. The first time you open the editor from the app, it uses the app's current theme; after you switch inside the editor, the choice is saved in the browser's localStorage. The CV page itself stays white because that is what is printed and exported.

The **Export** dialog offers two formats:

- **PDF** is an exact visual copy of each page.
- **DOCX (Word)** writes the canvas text into a single-column Word file that ATS parsers can read; freeform positioning is not kept in that file.

The CV/PDF is processed inside the browser and is not uploaded to a server. Scanned image-only PDFs have no selectable text, so the editor cannot import them. Basic touch editing works on mobile, but detailed layout work is easier on a desktop or laptop. The editor needs an internet connection to load its pinned browser libraries.

Freeform columns, icons, and decorative elements can weaken ATS parsing. Keep a simple one-column result for MNC or job-portal applications. For maximum ATS safety, use the structured **Existing Templates** flow.

## Clone from GitHub and run

Install [Python 3.10+](https://www.python.org/downloads/) first. On Windows, tick **Add Python to PATH** in the installer. Git is also required.

**Windows (PowerShell / CMD):**

```powershell
git clone https://github.com/satyamsingh24/My_Agent.git
cd My_Agent
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

After cloning, you can also double-click `RUN_MY_AGENT.bat` — it installs packages and starts the app.

**Linux / macOS:**

```bash
git clone https://github.com/satyamsingh24/My_Agent.git
cd My_Agent
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

The terminal prints a Local URL — open that in a browser. Unlock the lock screen with the app PIN. Stop the app with `Ctrl+C` in the terminal.

Optional (recommended): create a virtual environment after cloning so packages do not touch system Python:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

On Linux/macOS, activate with `source .venv/bin/activate`.

Download file names are random: `My_Agent_12345.pdf`, `.docx`, `.xlsx` (all three formats from one generation share the same number).

## Online demo (no code download)

Live app: **https://satyamsingh24.github.io/My_Agent/**

This page uses [stlite](https://github.com/whitphx/stlite) to run Streamlit on WebAssembly (Pyodide) **inside the visitor's browser**. There is no server and no install — open the link and unlock the lock screen to use the full frontend.

The first load downloads the Python runtime, so it can take 30–60 seconds and the app is a little slower than a local run. For daily use, the local commands above are still best.

For a faster hosted version, use **Streamlit Community Cloud** (free):

1. Open [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Choose **Create app** → **Deploy a public app from GitHub**.
3. Repository: `satyamsingh24/My_Agent`, Branch: `main`, Main file: `app.py`.
4. Press **Deploy**. The first build takes 2–3 minutes.

In hosted mode the app detects that it is running on a server or in the browser:

- An uploaded CV stays in that visitor's **browser session**, is not saved on the server disk, and is not visible to other visitors.
- Generated PDF/DOCX/XLSX files are available only from the download buttons; no copy is written on the server.
- The lock PIN lives in the public repository, so a hosted app should not be treated as private. It is fine as a demo.

## Where files are saved (local run)

- Existing uploaded CV: `app_data/existing_cv.pdf`
- Extracted information: `app_data/existing_cv.txt`
- CV metadata: `app_data/existing_cv.json`
- Generated CV: `applications/generated/`

## On another laptop

Copy the whole `My_Agent` folder by USB, ZIP, or Drive. Install Python and double-click `RUN_MY_AGENT.bat`. Nothing else needs to be installed.

## Important limitation

Ollama is only reachable at `localhost:11434` during a local Streamlit run. On Streamlit Cloud, configure Gemini or Groq. The GitHub Pages stlite/browser build cannot reliably reach local Ollama because of browser networking and CORS. Deterministic ATS matching and DOCX generation still work when AI is unavailable.
