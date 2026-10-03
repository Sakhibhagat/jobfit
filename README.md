# JobFit

An AI-powered resume tailoring tool. Upload your resume as a PDF and a job description, and JobFit analyzes your fit for the role, then rewrites your resume to better match what the job is asking for — matching your original resume's font and color so the output doesn't look like a generic template.

**Live demo:** https://jobfit-u1t9.onrender.com
*(Hosted on a free tier — may take 30-60 seconds to load on first visit if inactive.)*

## How it works

1. Upload your resume (PDF) and paste in the job description
2. The app extracts the resume's text, along with its actual font and heading color
3. An AI model (via the Groq API) analyzes your fit for the role and returns feedback
4. A second AI call rewrites your resume, tailored to the job
5. Preview the result, request further changes if needed, and download it as a Word document

## Tech stack

- **Backend:** Python, Flask
- **AI:** Groq API (GPT-OSS-120B)
- **PDF processing:** pypdf (text extraction), PyMuPDF (font/color extraction)
- **Document generation:** python-docx
- **Frontend:** HTML, CSS, JavaScript (no framework)
- **Deployment:** Render

## Setup

1. Clone this repo
2. Create a virtual environment: `python -m venv venv`
3. Activate it:
   - Windows: `venv\Scripts\activate`
   - Mac/Linux: `source venv/bin/activate`
4. Install dependencies: `pip install -r requirements.txt`
5. Create a `.env` file with your Groq API key: