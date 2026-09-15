import os
from dotenv import load_dotenv
from groq import Groq
from pypdf import PdfReader

# Load API key from .env file
load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

def analyze_resume(resume_text):
    prompt = f"""You are an expert resume reviewer and career coach. 
Analyze the following resume and provide:
1. Overall impression (2-3 sentences)
2. Top 3 strengths
3. Top 3 areas for improvement (be specific and actionable)
4. Any formatting or clarity issues
5. A revised version of one weak bullet point as an example

Resume:
{resume_text}
"""

    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )

    return response.choices[0].message.content

if __name__ == "__main__":
    pdf_path = input("Enter the path to your resume PDF: ").strip('"').strip("'")
    resume_text = extract_text_from_pdf(pdf_path)
    print("\nAnalyzing your resume...\n")
    feedback = analyze_resume(resume_text)
    print(feedback)