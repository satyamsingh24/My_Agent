# My_AGENT — ATS CV Builder

Standalone browser application — sirf Python chahiye, koi extra tool nahi.

## Access

App khulte hi PIN popup aata hai. PIN: **6932** (hardcoded in `app.py` as `APP_PIN`).

## Features

0. **Dark / Light mode** toggle top-right, aur har screen par **Exit to Front Page** button.
1. **Upload CV** — koi bhi selectable-text PDF upload karo (DevOps, Python, Java, etc.).
2. CV text local files mein save hota hai.
3. **Existing CV** — saved CV continue ya change karo.
4. Complete JD paste karo.
5. App CV aur JD mein common verified skills detect karti hai.
6. Attractive ATS-safe one-column layout: readable Arial/Helvetica fonts, blue
   section headings, clean spacing, no columns/icons/tables.
7. Download popup se **PDF**, **DOCX**, ya **XLSX** choose karo.
8. **Make your CV** — bina purana CV ke, form (name, email, number, address) aur
   ek details box se naya ATS CV banao. Details box mein skills, experience,
   education, address ya company ka JD — kuch bhi daal sakte ho.
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

## Files kahan save hoti hain

- Existing uploaded CV: `app_data/existing_cv.pdf`
- Extracted information: `app_data/existing_cv.txt`
- CV metadata: `app_data/existing_cv.json`
- Generated CV: `applications/generated/`

## Dusre laptop par

Poora `My_Agent` folder USB, ZIP, ya Drive se copy karo. Python install karo aur
`RUN_MY_AGENT.bat` double-click karo. Iske alawa kuch install karne ki zaroorat
nahi hai.

## Important limitation

Yeh offline deterministic app hai, generative AI nahi. Yeh uploaded CV ko readable
ATS PDF banati hai aur JD/CV ke truthful matching keywords prioritize karti hai.
Deep sentence rewriting ke liye external AI API ki zaroorat hogi.
