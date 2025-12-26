# import streamlit as st
# import smtplib
# import openai
# import fitz  # PyMuPDF
# import os
# from email.message import EmailMessage

# # ---------------- CONFIG ----------------
# SENDER_EMAIL = "sksandy3869@gmail.com"  # Hardcoded
# RESUME_PATH = "sandeep__resume.pdf"
# OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")  # Set in terminal

# # ---------------- STREAMLIT UI ----------------
# st.set_page_config(page_title="AI Job Mailer", page_icon="📧")
# st.title("📧 AI Job Application Mailer")
# st.markdown("Minimal input: only Recruiter Email and Position Name. Everything else is automatic.")

# # Inputs
# recipient_email = st.text_input("Recruiter / Careers Email")
# position = st.text_input("Position Name (e.g. Data Analyst, BI Analyst)")

# # Optional company input (only used if recipient email is generic Gmail)
# company_override = st.text_input("Company Name (Optional, if Gmail or generic email)")

# # ---------------- HELPERS ----------------
# def extract_company_from_email(email):
#     try:
#         domain = email.split("@")[1]
#         if domain.lower() == "gmail.com":
#             return company_override if company_override else "Company"
#         company = domain.split(".")[0]
#         return company.replace("-", " ").title()
#     except:
#         return company_override if company_override else "Company"

# def read_resume(pdf_path):
#     doc = fitz.open(pdf_path)
#     text = ""
#     for page in doc:
#         text += page.get_text()
#     return text.strip()

# def generate_email_body(company, position, resume_text):
#     prompt = f"""
# You are a professional job applicant.

# Using the resume below, write a SHORT, recruiter-friendly job application email.

# Rules:
# - 5–6 lines max
# - Formal, confident, concise
# - No emojis
# - Mention the position and company naturally

# Company: {company}
# Position: {position}

# Resume:
# {resume_text}
# """
#     response = openai.ChatCompletion.create(
#         model="gpt-4o-mini",
#         messages=[{"role": "user", "content": prompt}],
#         temperature=0.3
#     )
#     return response.choices[0].message.content.strip()

# def generate_email_fallback(company, position):
#     subject = f"Application for {position}"
#     body = f"""Dear Hiring Manager,

# I am writing to apply for the {position} position at {company}.

# I bring experience in data analysis and reporting, with a focus on transforming data into actionable insights that support informed decision-making. I am eager to contribute to a structured, data-driven organization like {company}.

# I would welcome the opportunity to discuss my fit for this role.

# Here is my portfolio link: https://sandeepkumar3869.github.io/PORTFOLIO/

# Best regards,
# Sandeep Kumar
# """
#     return subject, body

# def send_email(sender, app_password, to_email, subject, body, resume_path):
#     msg = EmailMessage()
#     msg["From"] = sender
#     msg["To"] = to_email
#     msg["Subject"] = subject
#     msg.set_content(body)

#     with open(resume_path, "rb") as f:
#         msg.add_attachment(
#             f.read(),
#             maintype="application",
#             subtype="pdf",
#             filename=os.path.basename(resume_path)
#         )

#     with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
#         server.login(sender, app_password)
#         server.send_message(msg)

# # ---------------- ACTION ----------------
# if st.button("🚀 Generate & Send Email"):
#     if not recipient_email or not position:
#         st.error("Please enter both Recruiter Email and Position Name.")
#     else:
#         try:
#             openai.api_key = OPENAI_API_KEY
#             app_password = os.getenv("GMAIL_APP_PASSWORD")
#             if not app_password:
#                 st.error("Gmail App Password not found in terminal environment. Set using: export GMAIL_APP_PASSWORD='your_app_password'")
#             else:
#                 company = extract_company_from_email(recipient_email)
#                 resume_text = read_resume(RESUME_PATH)

#                 # AI email generation with fallback
#                 try:
#                     email_body = generate_email_body(company, position, resume_text)
#                     st.success("AI-generated email created.")
#                 except Exception:
#                     st.warning("AI quota exceeded. Using safe fallback template.")
#                     subject, email_body = generate_email_fallback(company, position)

#                 subject = f"Application for {position} – {company}"

#                 send_email(
#                     sender=SENDER_EMAIL,
#                     app_password=app_password,
#                     to_email=recipient_email,
#                     subject=subject,
#                     body=email_body,
#                     resume_path=RESUME_PATH
#                 )

#                 st.success("✅ Email sent successfully.")
#                 st.text_area("📄 Email Preview", email_body, height=250)

#         except Exception as e:
#             st.error(f"❌ Error: {str(e)}")

from datetime import datetime
print("App started at:", datetime.utcnow().isoformat())
import os
import time
import smtplib
import pandas as pd
import openai
import fitz  # PyMuPDF
import gspread
from email.message import EmailMessage
from google.oauth2.service_account import Credentials
from datetime import *
import pytz
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

            # try:
            #     body = generate_email_body(company, post_name, resume_text)
            # except Exception:
            #     body = generate_email_fallback(company, post_name)

            # subject = f"Application for {post_name.title()} – {company}"

            # send_email(mail_id, subject, body)

            # sheet.update_cell(idx + 2, 4, body)     # body
            # sheet.update_cell(idx + 2, 5, "DONE")   # status
            try:
                body = generate_email_body(company, post_name, resume_text)
            except Exception:
                body = generate_email_fallback(company, post_name)

            # Body_TimeStamp = datetime.now().strftime("%m/%d/%Y %H:%M:%S")
            

            ist = pytz.timezone("Asia/Kolkata")
            Body_TimeStamp = datetime.now(ist).strftime("%m/%d/%Y %H:%M:%S")

            subject = f"Application for {post_name.title()} – {company}"

            send_email(mail_id, subject, body)

            sheet.update_cell(idx + 2, 4, body)              # body
            sheet.update_cell(idx + 2, 5, "DONE")            # status
            sheet.update_cell(idx + 2, 7, Body_TimeStamp)    # Body_TimeStamp


        except Exception as e:
            sheet.update_cell(idx + 2, 4, str(e))
            sheet.update_cell(idx + 2, 5, "FAILED")

        time.sleep(10)  # rate limiting (important)

    print("Batch completed successfully.")

# ---------------- RUN ----------------
if __name__ == "__main__":
    main()