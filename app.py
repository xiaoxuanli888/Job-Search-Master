import io
import re
import zipfile
from datetime import datetime
from typing import Dict, Any, List, Optional

import streamlit as st
from docx import Document
from openai import OpenAI
from pydantic import BaseModel, Field


# -----------------------------
# Helpers
# -----------------------------
def load_bytes(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def load_text(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def safe_filename(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]+", "_", (s or "").strip())
    return s.strip("_") or "output"


def replace_tokens_in_doc(doc: Document, mapping: Dict[str, str]) -> None:
    """
    Replace {{TOKEN}} in paragraphs AND tables.
    Best practice: keep placeholders in their own paragraph or typed in one go.
    """
    def replace_in_paragraph(paragraph):
        if "{{" not in paragraph.text:
            return
        for run in paragraph.runs:
            for k, v in mapping.items():
                token = f"{{{{{k}}}}}"
                if token in run.text:
                    run.text = run.text.replace(token, v)

    for p in doc.paragraphs:
        replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)


def doc_contains_token(doc: Document, token: str) -> bool:
    needle = f"{{{{{token}}}}}"

    for p in doc.paragraphs:
        if needle in p.text:
            return True

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if needle in p.text:
                        return True

    return False


def make_zip(cv_bytes: bytes, cl_bytes: bytes, cv_name: str, cl_name: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(cv_name, cv_bytes)
        z.writestr(cl_name, cl_bytes)
    return buf.getvalue()


# -----------------------------
# Model output schema (Pydantic)
# We request 6 Versuni bullets so we can "add one extra bullet point" when needed.
# -----------------------------
class CVCoverOutput(BaseModel):
    company: str
    role_title: str
    about_me: str

    # 6 bullets: lets us add an extra bullet point when fixing unsupported claims
    versuni_bullets: List[str] = Field(min_length=6, max_length=6)

    philips_bullets: List[str] = Field(min_length=2, max_length=2)
    cover_letter_body: str

    # If anything is not supported by base materials, list it here
    new_claims_not_in_base: List[str] = Field(default_factory=list)


# -----------------------------
# OpenAI call (Responses API + parse)
# -----------------------------
def generate_structured(
    job_description: str,
    prompt_rules: str,
    base_cv: str,
    base_cover_letter: str,
    model: str,
    fix_mode: bool = False,
    issues_to_fix: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    fix_mode=False: normal generation
    fix_mode=True: second pass if the model flagged unsupported claims.
                   It must remove/adjust those claims AND add 1 extra CV bullet (already enforced by schema).
    """
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

    issues_text = ""
    if fix_mode and issues_to_fix:
        issues_text = "\n".join([f"- {x}" for x in issues_to_fix])

    system_instructions = f"""
You are a CV + cover letter customization engine.

You are given:
1) The user's ORIGINAL CV content (base CV)
2) The user's ORIGINAL cover letter (base cover letter)
3) A job description

Your job:
- Rewrite ONLY: ABOUT_ME, Versuni bullets (exactly 6), Philips bullets (2), and the cover letter body.
- The output must be high quality, concrete, ATS-friendly, and aligned with the job description.
- All content MUST be strictly grounded in the base CV + base cover letter. Do NOT invent employers, titles, dates, metrics, tools, budgets, languages, or results.
- Preserve the user's voice from the base cover letter.
- Output MUST match the schema exactly.

IMPORTANT SAFETY CHECK:
- If you add any claim not clearly supported by the base CV or base cover letter,
  list it in new_claims_not_in_base. Otherwise return [].

SECOND-PASS FIX MODE (only if requested):
- If fix_mode is enabled, you must eliminate or reword the unsupported claims listed below so that new_claims_not_in_base becomes [].
- Also ensure the 6th Versuni bullet is an additional strong bullet, derived by rewording/recombining what is already in the base materials and aligned to the job description.

Unsupported claims to fix (if any):
{issues_text if issues_text else "(none)"}

BASE CV (source of truth):
{base_cv}

BASE COVER LETTER (source of truth):
{base_cover_letter}

USER PROMPT / RULES:
{prompt_rules}
""".strip()

    user_msg = f"JOB DESCRIPTION:\n{job_description}"
    if fix_mode:
        user_msg += "\n\nPlease revise the draft to remove unsupported claims and keep new_claims_not_in_base empty, while maintaining strong impact and including 6 Versuni bullets."

    resp = client.responses.parse(
        model=model,
        input=[
            {"role": "system", "content": system_instructions},
            {"role": "user", "content": user_msg},
        ],
        text_format=CVCoverOutput,
    )

    return resp.output_parsed.model_dump()


# -----------------------------
# Build docs from templates
# -----------------------------
def build_docs(
    data: Dict[str, Any],
    cv_template_bytes: bytes,
    cl_template_bytes: bytes,
) -> tuple[bytes, bytes]:
    cv_doc = Document(io.BytesIO(cv_template_bytes))
    cl_doc = Document(io.BytesIO(cl_template_bytes))

    # Some templates may not have {{VERSUNI_BULLET_6}}.
    has_bullet_6 = doc_contains_token(cv_doc, "VERSUNI_BULLET_6")

    # If no placeholder for bullet 6, append it to bullet 5 so you still get the "extra bullet".
    bullet5 = data["versuni_bullets"][4]
    bullet6 = data["versuni_bullets"][5]
    if not has_bullet_6 and bullet6.strip():
        bullet5 = f"{bullet5}\n• {bullet6}"

    mapping = {
        "ROLE_TITLE": data["role_title"],
        "COMPANY": data["company"],
        "ABOUT_ME": data["about_me"],
        "VERSUNI_BULLET_1": data["versuni_bullets"][0],
        "VERSUNI_BULLET_2": data["versuni_bullets"][1],
        "VERSUNI_BULLET_3": data["versuni_bullets"][2],
        "VERSUNI_BULLET_4": data["versuni_bullets"][3],
        "VERSUNI_BULLET_5": bullet5,
        "VERSUNI_BULLET_6": bullet6 if has_bullet_6 else "",
        "PHILIPS_BULLET_1": data["philips_bullets"][0],
        "PHILIPS_BULLET_2": data["philips_bullets"][1],
        "DATE_TODAY": datetime.now().strftime("%d %B %Y, Berlin"),
        "COVER_LETTER_BODY": data["cover_letter_body"],
    }

    replace_tokens_in_doc(cv_doc, mapping)
    replace_tokens_in_doc(cl_doc, mapping)

    cv_out = io.BytesIO()
    cl_out = io.BytesIO()
    cv_doc.save(cv_out)
    cl_doc.save(cl_out)
    return cv_out.getvalue(), cl_out.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="CV + Cover Letter Agent", layout="wide")
st.title("CV + Cover Letter Agent")

model = st.secrets.get("MODEL", "gpt-5.2")

# Load templates
CV_TEMPLATE_CORPORATE = "templates/cv_template.docx"
CV_TEMPLATE_STARTUP = "templates/cv_template_startup.docx"
CL_TEMPLATE = "templates/cover_letter_template.docx"

# Controls row
colA, colB, colC = st.columns([1, 1, 1])
with colA:
    template_choice = st.radio(
        "CV template",
        options=["Corporate", "Startup"],
        index=0,
        horizontal=True,
    )
with colB:
    show_debug = st.checkbox("Show debug JSON", value=False)
with colC:
    use_test = st.checkbox("Use test job description", value=False)

cv_template_path = CV_TEMPLATE_CORPORATE if template_choice == "Corporate" else CV_TEMPLATE_STARTUP

cv_template_bytes = load_bytes(cv_template_path)
cl_template_bytes = load_bytes(CL_TEMPLATE)

# Base materials
prompt_rules = load_text("prompt/instructions.txt")
base_cv = load_text("prompt/base_cv.txt")
base_cover_letter = load_text("prompt/base_cover_letter.txt")

# Optional test JD
test_jd_path = "prompt/test_job_description.txt"
try:
    test_jd = load_text(test_jd_path)
except FileNotFoundError:
    test_jd = ""

st.caption("Paste a job description → Generate → Download a ZIP with CV + cover letter (.docx).")

default_text = test_jd if (use_test and test_jd.strip()) else ""
job_description = st.text_area("Job description", value=default_text, height=320)

if st.button("Generate", type="primary"):
    if not job_description.strip():
        st.error("Please paste a job description.")
        st.stop()

    # Placeholder diagnostics
    cv_doc_check = Document(io.BytesIO(cv_template_bytes))
    cl_doc_check = Document(io.BytesIO(cl_template_bytes))

    required_cv_tokens = [
        "ABOUT_ME",
        "VERSUNI_BULLET_1", "VERSUNI_BULLET_2", "VERSUNI_BULLET_3", "VERSUNI_BULLET_4", "VERSUNI_BULLET_5",
        # Bullet 6 is optional in the template; we handle it gracefully.
        "PHILIPS_BULLET_1", "PHILIPS_BULLET_2",
    ]
    required_cl_tokens = ["COMPANY", "COVER_LETTER_BODY"]

    missing = []
    for t in required_cv_tokens:
        if not doc_contains_token(cv_doc_check, t):
            missing.append(f"CV template missing {{{{{t}}}}}")
    for t in required_cl_tokens:
        if not doc_contains_token(cl_doc_check, t):
            missing.append(f"Cover letter template missing {{{{{t}}}}}")

    if missing:
        st.error("Your templates are missing required placeholders:")
        for m in missing:
            st.write(f"- {m}")
        st.stop()

    # 1) First pass
    with st.spinner("Generating..."):
        data = generate_structured(
            job_description=job_description,
            prompt_rules=prompt_rules,
            base_cv=base_cv,
            base_cover_letter=base_cover_letter,
            model=model,
            fix_mode=False,
        )

    # 2) If unsupported claims exist, auto-fix once and add extra bullet via schema (6 bullets always)
    if data.get("new_claims_not_in_base"):
        with st.spinner("Fixing unsupported claims and strengthening CV..."):
            data_fixed = generate_structured(
                job_description=job_description,
                prompt_rules=prompt_rules,
                base_cv=base_cv,
                base_cover_letter=base_cover_letter,
                model=model,
                fix_mode=True,
                issues_to_fix=data["new_claims_not_in_base"],
            )

        # If still problematic after one auto-fix, stop and show issues
        if data_fixed.get("new_claims_not_in_base"):
            st.error("I tried to fix unsupported claims, but some remain. Please review:")
            st.write(data_fixed["new_claims_not_in_base"])
            if show_debug:
                st.subheader("Debug: AI JSON output (after fix attempt)")
                st.json(data_fixed)
            st.stop()

        data = data_fixed

    # Build docs + ZIP
    cv_bytes, cl_bytes = build_docs(data, cv_template_bytes, cl_template_bytes)

    company = safe_filename(data["company"])
    role = safe_filename(data["role_title"])

    cv_name = f"Xiaoxuan_Li_CV_{company}_{role}.docx"
    cl_name = f"Xiaoxuan_Li_Cover_Letter_{company}_{role}.docx"
    zip_name = f"Xiaoxuan_Li_{company}_{role}.zip"

    zip_bytes = make_zip(cv_bytes, cl_bytes, cv_name, cl_name)

    st.success("Done!")
    st.download_button(
        "Download CV + Cover Letter (ZIP)",
        data=zip_bytes,
        file_name=zip_name,
        mime="application/zip",
    )

    if show_debug:
        st.subheader("Debug: AI JSON output")
        st.json(data)
