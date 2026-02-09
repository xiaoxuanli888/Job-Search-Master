# 🌟 AI-Powered CV & Cover Letter Customization Agent

### Make every job application faster, smarter, and perfectly tailored.

This project is my personal AI agent designed to **generate job-specific CVs and cover letters** in seconds.
I built it because tailoring applications manually is time-consuming, exhausting, and often inconsistent.

Now the agent does the heavy lifting. 

This generated CV and cover letters are based on my original CV and cover letters.

I make sure that the agent only helps to adjust the wordings, all the facts, achievements and experience are grounded in real experience. 

No Hallucinations!

---

## ✨ What This AI Agent Does

### 🔹 1. Reads and understands any job description

I paste a job description into the app, and the agent breaks it down into:

* key requirements
* responsibilities
* skills
* priority ordering

This ensures the generated CV and cover letter speak the same language as the employer.

---

### 🔹 2. Rewrites my CV content specifically for *that* job

The AI restructures and rewrites:

* **About Me** section
* **Versuni experience bullets** (6 bullets, reordered to match JD priorities)
* **Philips experience bullets** (2 bullets)

It aligns each bullet point to the job requirements so recruiters can instantly see the match.

---

### 🔹 3. Generates a customized cover letter

The agent produces a personalized cover letter with:

* tailored opening
* JD-aligned bullet points
* country-specific tone (Germany vs. Netherlands)
* role-specific emphasis (corporate vs. startup)

It mirrors the job description order, which makes scanning extremely easy for hiring managers.

---

### 🔹 4. Uses real CV content — no hallucinations

This is not a “creative writing” AI. It is grounded by design.

The model **cannot** invent:

* companies
* job titles
* dates
* languages
* tools
* responsibilities
* metrics

If the AI mistakenly adds something not present in my original CV, it automatically detects it and rewrites the content in a second pass until everything is factual.

---

### 🔹 5. Uses Word templates for formatting consistency

Instead of letting AI freestyle formats, I created modular templates for:

#### CV templates

* Germany — Corporate
* Germany — Startup
* Netherlands — Corporate
* Netherlands — Startup

#### Cover letter templates

* Germany
* Netherlands

The AI only fills placeholders (like `{{ABOUT_ME}}` or `{{VERSUNI_BULLET_1}}`) so the layout remains perfect every time.

---

### 🔹 6. Outputs downloadable files automatically

With one click, I get a ZIP file containing:

* A fully formatted **CV.docx**
* A fully formatted **CoverLetter.docx**

Both tailored to the job description.

---

## 🧠 How It Works (Simple Explanation)

1. I select:

   * Country (DE / NL)
   * Style (Corporate / Startup)

2. I paste a job description.

3. The agent loads:

   * my base CV text
   * my base cover letter
   * the job description
   * my chosen templates

4. It sends those inputs to OpenAI with strict instructions and a strict schema.

5. AI rewrites only the allowed sections using my real content.

6. The system checks for any invented claims:

   * If found → AI is forced to rewrite (auto-fix mode)
   * If clean → proceed

7. The rewritten content is injected into the Word templates.

8. The downloadable ZIP is created.

---

## 📁 Project Structure

```
cv-agent/
│
├── app.py                     # Main Streamlit application
├── requirements.txt           # Dependencies
│
├── prompt/
│   ├── base_cv.txt
│   ├── base_cover_letter.txt
│   ├── instructions.txt
│   └── test_job_description.txt
│
└── templates/
    ├── cv_template_DE_corporate.docx
    ├── cv_template_DE_startup.docx
    ├── cv_template_NL_corporate.docx
    ├── cv_template_NL_startup.docx
    ├── cover_letter_template_DE.docx
    └── cover_letter_template_NL.docx
```

---

## 🛠️ Why I Built This

Because job searching is emotionally draining — and repetitive.
Every role requires a fresh CV and cover letter, but the core content is always the same.

The agent helps me:

* save hours of rewriting
* stay consistent
* reduce stress
* highlight the right things
* avoid overselling or inventing things
* adapt quickly when switching markets (DE vs NL)

It gives me more time to focus on preparing for interviews and less time formatting documents.

---

## 🚀 Technologies Used

* **Python**
* **Streamlit**
* **OpenAI Responses API**
* **Pydantic** for schema validation
* **python-docx** for template filling
* **ZIP generation** for downloads

---

## 🔐 Security and Privacy

* API key is stored in `.streamlit/secrets.toml` (never committed to GitHub)
* Base CV and cover letter are loaded locally or from private repo
* No user data is stored on the server
* GitHub does NOT contain any sensitive keys

---

## 🧭 Future Improvements

Planned next steps:

* Bold key phrases in bullets automatically
* Add recruiter outreach email generator
* Add PDF export option
* Add AI interview question generator based on the JD

---

## ❤️ Contributions & Feedback

This is a personal project, but I welcome:

* suggestions
* feature requests
* pull requests

Feel free to open an issue if you want a feature or run into trouble.

