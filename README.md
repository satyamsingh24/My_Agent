# My_AGENT — ATS CV Builder

Standalone browser application — sirf Python chahiye, koi extra tool nahi.

## Access

App khulte hi PIN popup aata hai. PIN: **6932** (hardcoded in `app.py` as `APP_PIN`).

## Features

0. **Dark / Light mode** toggle top-right, aur har screen par **Exit to Front Page** button.
1. **Upload CV** — koi bhi selectable-text PDF upload karo (DevOps, Python, Java, etc.).
2. CV text local files mein save hota hai.
3. **Existing CV** — saved CV continue/change karo aur Upload CV wala complete
   JD + ATS score + AI/fallback + truthful rebuild workflow use karo.
4. Complete JD paste karo.
5. App CV aur JD mein common verified skills detect karti hai.
6. Attractive ATS-safe one-column layout: readable Arial/Helvetica fonts, blue
   section headings, clean spacing, no columns/icons/tables.
7. Download popup se **PDF**, **DOCX**, ya **XLSX** choose karo.
8. **Make your CV** — bina purana CV ke form se naya CV banao, target JD paste
   karo, aur wahi complete ATS score + AI/fallback + truthful confirmation
   workflow use karo. Output PDF, DOCX aur XLSX mein milta hai.
9. Reference resume sirf **layout aur heading order** ke liye use hota hai —
   content sirf tumhari di hui details ka rehta hai, kisi doosre CV se nahi.
10. Kam details do to bhi CV professional banta hai: role + skills se
    professional summary, grouped SKILLS (Cloud / DevOps / Databases...) aur
    CORE COMPETENCIES bullets tumhare hi keywords se bante hain.
11. Form ki details badalne par warning aati hai — dobara **Make CV** dabana
    zaroori hai, warna purana CV dikhta rahega.
12. Front page par professional animated background chalta hai (moving gradient
    glow + slow grid + hero sheen) — pure CSS, koi internet ya extra file nahi.
    Apni image chahiye to `assets/background.jpg` (ya `.png` / `.webp`) rakh do;
    app usko dim + slow-zoom karke backdrop bana degi. File hatane par CSS
    animation wapas.
13. Sidebar ke top par **My Profile** page hai: owner ki passport-size photo
    (`assets/owner_photo.jpg`) aur basic details code ke saath ship hoti hain,
    aur **My CV for Reference** button reference resume download karta hai.
14. Left sidebar mein **Cold Mail for Referral** CV ko read karke short HR email
    banata hai. Fresher, internship-only aur experienced profiles ke liye alag
    truthful wording use hoti hai; CV ke bahar ki skill/experience add nahi hoti.
15. Sidebar ke **Settings** page par Log out button session lock karke PIN screen
    par wapas le jaata hai.
16. Top-right **AI** button password `6932AI` se provider settings kholta hai.
    Ollama local-first hai; failure par configured Gemini aur Groq free-tier
    providers try ho sakte hain. Cloud API keys sirf current session memory mein
    rehti hain aur logout par clear ho jaati hain.
17. **Templates → Edit Your CV** ek browser-only freeform editor kholta hai.
    Selectable-text PDF import karke text boxes move/resize/rotate/edit karo,
    fonts, colours, shapes, lines aur images use karo, 10 template previews mein
    se koi design apply karo, undo/redo chalao aur PDF ya DOCX export karo.

Har generated CV `profile/Satyam_Dev_Resume_ATS.pdf` ke reference layout ko follow
karta hai: Name → Role → Contact → SUMMARY → SKILLS → JD-MATCHED SKILLS →
PROFESSIONAL EXPERIENCE → INTERNSHIPS → PROJECTS → EDUCATION → CERTIFICATIONS.
Declaration/References jaise ATS-noise blocks hata diye jaate hain.

PDF aur DOCX ATS-friendly hain. XLSX sirf reference/editing ke liye hai; Excel
CV job portal par upload karna recommended nahi hai.

Skill matching DevOps/cloud ke saath Java full stack, Python, frontend,
backend, mobile, data/AI, testing aur security technologies cover karti hai.

App skills ya experience invent nahi karti. JD mein jo requirement hai lekin uploaded
CV mein verify nahi hoti, woh missing list mein dikhayi jaati hai.

## Free AI providers

Upload CV flow deterministic ATS checks ke baad optional AI wording suggestions
use karta hai. Default provider chain:

1. **Ollama** — free/local, CV computer se bahar nahi jaata.
2. **Gemini** — free tier; API key AI settings dialog mein enter karo.
3. **Groq** — free tier; API key AI settings dialog mein enter karo.

Ollama ke liye [Ollama](https://ollama.com/) install karke model pull aur server
start karo:

```powershell
ollama pull llama3.2:3b
ollama serve
```

Provider settings ka password: **6932AI**. Yeh password public source code mein
hardcoded hai, isliye real security boundary nahi hai. Gemini/Groq select ya
fallback hone par uploaded CV aur JD us cloud provider ko bheje jaate hain.
Missing JD skills AI khud add nahi karti; user ko pehle unhe truthful confirmation
list mein tick karna padta hai.

## Freeform CV editor

Home page par **4. Templates** kholo:

1. **Existing Templates** structured, JD-aligned aur ATS-safe CV flow kholta hai.
2. **Edit Your CV** visual editor kholta hai. **Import PDF** se selectable-text
   CV do; phir labelled toolbar ke text, insert, arrange, history, zoom aur page
   tools use karo aur **Export** se format chuno.

Left panel har template ka actual page preview image dikhata hai; kisi preview par
click karne par imported CV ki information usi design mein editable objects ke roop
mein rebuild ho jaati hai. Template switch sirf layout/style badalta hai; app nayi
skill ya experience invent nahi karti. Preview images `static/editor/previews/`
mein hain aur `CV_TEMPLATES` badalne ke baad
`python -m scripts.export_template_previews` se regenerate hoti hain.

Editor ke header mein apna dark/light switch hai. Pehli baar app se editor kholne
par wahi theme use hoti hai jo app par selected hai; editor mein switch dabate hi
choice browser ke localStorage mein save ho jaati hai. CV page hamesha white rehta
hai kyunki wahi print/export hota hai.

**Export** dialog do formats deta hai:

- **PDF** har page ki exact visual copy deta hai.
- **DOCX (Word)** canvas ke text ko single-column Word file mein likhta hai, jise
  ATS aasani se parse karta hai; freeform positioning is file mein nahi rehti.

CV/PDF browser ke andar process hota hai aur server par upload nahi hota. Scanned
image-only PDFs mein selectable text nahi hota, isliye unhe editor import nahi
kar sakta. Mobile par basic touch editing supported hai, lekin detailed layout
editing desktop/laptop par easier hai. Editor ki pinned browser libraries load
karne ke liye internet connection chahiye.

Freeform columns, icons aur decorative elements ATS parsing ko weak kar sakte
hain. MNC/job-portal application ke liye simple one-column result rakho; maximum
ATS safety ke liye **Existing Templates** wala structured flow use karo.

## GitHub se clone karke run

Pehle [Python 3.10+](https://www.python.org/downloads/) install karo. Windows par installer mein **Add Python to PATH** tick karo. Git bhi chahiye.

**Windows (PowerShell / CMD):**

```powershell
git clone https://github.com/satyamsingh24/My_Agent.git
cd My_Agent
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Ya clone ke baad `RUN_MY_AGENT.bat` par double-click — yeh packages install karke app start kar dega.

**Linux / macOS:**

```bash
git clone https://github.com/satyamsingh24/My_Agent.git
cd My_Agent
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

Terminal mein URL dikhega, jaise `http://localhost:6932` — browser mein wahi kholo. PIN: **6932**. Band karne ke liye terminal mein `Ctrl+C`.

Optional (recommended): clone ke baad virtual environment banao, taaki packages system Python ko na chhedein:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Linux/macOS par activate: `source .venv/bin/activate`

Download file names random hote hain: `My_Agent_12345.pdf`, `.docx`, `.xlsx`
(ek generation ke teeno format same number use karte hain).

## Online demo (bina code download kiye)

Live app: **https://satyamsingh24.github.io/My_Agent/**

Yeh page [stlite](https://github.com/whitphx/stlite) use karta hai, jo Streamlit
ko WebAssembly (Pyodide) par visitor ke **browser ke andar** chalata hai. Koi
server nahi, koi install nahi — link kholo aur PIN `6932` daal kar poora
frontend use karo.

Dhyan rahe: pehli baar Python runtime download hota hai, isliye load hone mein
30-60 second lag sakte hain aur app local run se thodi slow chalti hai. Roz ke
use ke liye local run (upar wale commands) hi best hai.

Zyada tez hosted version chahiye to **Streamlit Community Cloud** (free) use karo:

1. [share.streamlit.io](https://share.streamlit.io) kholo aur GitHub se sign in karo.
2. **Create app** → **Deploy a public app from GitHub** choose karo.
3. Repository: `satyamsingh24/My_Agent`, Branch: `main`, Main file: `app.py`.
4. **Deploy** dabao. Pehli build mein 2-3 minute lagte hain.

Hosted mode mein app khud detect kar leti hai ki woh server/browser par chal rahi hai:

- Uploaded CV sirf us visitor ke **browser session** mein rehta hai, server ki
  disk par save nahi hota, aur doosre visitors ko dikhta nahi.
- Generated PDF/DOCX/XLSX sirf download button se milte hain; server par file
  copy nahi banti.
- PIN (`6932`) public repo mein hai, isliye hosted app ko private samajhna galat
  hoga — demo ke liye theek hai.

## Files kahan save hoti hain (local run)

- Existing uploaded CV: `app_data/existing_cv.pdf`
- Extracted information: `app_data/existing_cv.txt`
- CV metadata: `app_data/existing_cv.json`
- Generated CV: `applications/generated/`

## Dusre laptop par

Poora `My_Agent` folder USB, ZIP, ya Drive se copy karo. Python install karo aur
`RUN_MY_AGENT.bat` double-click karo. Iske alawa kuch install karne ki zaroorat
nahi hai.

## Important limitation

Ollama sirf local Streamlit run par `localhost:11434` se accessible hota hai.
Streamlit Cloud par Gemini/Groq configure karo. GitHub Pages ka stlite/browser
build browser networking aur CORS restrictions ki wajah se local Ollama access
reliably nahi kar sakta. AI unavailable ho tab bhi deterministic ATS matching aur
DOCX generation chalti rehti hai.
