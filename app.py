from datetime import datetime
print("App started at:", datetime.utcnow().isoformat())
import os
import smtplib

import pandas as pd
import openai
import fitz  # PyMuPDF
import gspread
from email.message import EmailMessage
from google.oauth2.service_account import Credentials


import time


# ---------------- CONFIG ----------------
SENDER_EMAIL = "sksandy3869@gmail.com"
RESUME_PATH = "sandeep__resume.pdf"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

SHEET_ID = "15vQWXea6m0PGJKEYU_tLfUVa5F2Bw7JTTMkBXweUi9g"
SHEET_NAME = "Sheet1"  # change only if your tab name is different

# ---------------- GOOGLE SHEETS ----------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def connect_sheet():
    creds = Credentials.from_service_account_file(
        "credentials.json",
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

# ---------------- HELPERS ----------------
def extract_company_from_email(email, company_from_sheet):
    """
    Rules:
    1. Generic emails (gmail, yahoo, etc.) → ALWAYS use sheet company name
    2. Corporate emails → extract from domain
    3. Absolute fallback → 'Company'
    """
    try:
        domain = email.split("@")[1].lower()

        if domain in ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com"]:
            return company_from_sheet.strip() if company_from_sheet else "Company"

        company = domain.split(".")[0]
        return company.replace("-", " ").title()

    except Exception:
        return company_from_sheet.strip() if company_from_sheet else "Company"

def read_resume(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text.strip()

def generate_email_body(company, position, resume_text):
    prompt = f"""
You are a professional job applicant.

Write a SHORT recruiter-friendly job application email.

Rules:
- 5–6 lines max
- Formal, concise
- No emojis
- Mention position and company naturally

Company: {company}
Position: {position}

Resume:
{resume_text}
"""
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )
    return response.choices[0].message.content.strip()

def generate_email_fallback(company, position):
    return f"""Dear Hiring Manager,

I am writing to apply for the {position} position at {company}.

I bring experience in data analysis and reporting, with a focus on transforming data into actionable insights. I am eager to contribute to a data-driven organization like {company}.

Portfolio: https://sandeepkumar3869.github.io/PORTFOLIO/

Best regards,  
Sandeep Kumar
"""

def send_email(to_email, subject, body):
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    with open(RESUME_PATH, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(RESUME_PATH),
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, GMAIL_APP_PASSWORD)
        server.send_message(msg)

# ---------------- MAIN EXECUTION ----------------
def main():
    if not OPENAI_API_KEY or not GMAIL_APP_PASSWORD:
        raise EnvironmentError("Missing OPENAI_API_KEY or GMAIL_APP_PASSWORD")

    openai.api_key = OPENAI_API_KEY

    sheet = connect_sheet()
    df = pd.DataFrame(sheet.get_all_records())

    resume_text = read_resume(RESUME_PATH)

    for idx, row in df.iterrows():
        # Skip already processed rows
        if str(row.get("status", "")).upper() == "DONE":
            continue

        try:
            mail_id = str(row.get("Mail_ID", "")).strip()
            post_name = str(row.get("Post_name", "")).strip()
            company_from_sheet = str(row.get("Company_name", "")).strip()

            if not mail_id or not post_name:
                raise ValueError("Missing Mail_ID or Post_name")
            

            company = extract_company_from_email(mail_id, company_from_sheet)

            try:
                body = generate_email_body(company, post_name, resume_text)
            except Exception:
                body = generate_email_fallback(company, post_name)

            subject = f"Application for {post_name.title()} – {company}"

            send_email(mail_id, subject, body)

            sheet.update_cell(idx + 2, 4, body)     # body
            sheet.update_cell(idx + 2, 5, "DONE")   # status
            


        except Exception as e:
            sheet.update_cell(idx + 2, 4, str(e))
            sheet.update_cell(idx + 2, 5, "FAILED")

        time.sleep(10)  # rate limiting (important)
        

    print("Batch completed successfully.")

# ---------------- RUN ----------------
if __name__ == "__main__":
    main()