import smtplib
import os
import re
from email.message import EmailMessage

# ----------------------------
# CONFIG
# ----------------------------
SENDER_EMAIL = "sksandy3869@gmail.com"
APP_PASSWORD = "rzjo sfuv tpxy goeq"
RESUME_PATH = "sandeep__resume.pdf"

# ----------------------------
# UTIL: Extract company name
# ----------------------------
def extract_company(email):
    domain = email.split("@")[-1]
    domain = domain.replace("www.", "")
    name_part = domain.split(".")[0]

    # remove common junk words
    junk = ["careers", "career", "jobs", "hr"]
    for j in junk:
        name_part = name_part.replace(j, "")

    # clean & format
    name_part = re.sub(r"[-_]", " ", name_part).strip()
    return name_part.title()

# ----------------------------
# UTIL: Draft email
# ----------------------------
def generate_email(company, role):
    subject = f"Application for {role}"

    body = f"""
Dear Hiring Manager,

I am writing to apply for the {role} position at {company}.

I bring a strong foundation in data analysis and reporting, with experience converting complex datasets into actionable business insights. I am eager to contribute to an organization that values precision, structure, and data-driven decision-making.

Please find my resume attached for your review. I would welcome the opportunity to discuss how my skills can support {company}'s continued success.

Best regards,  
Sandeep Kumar
"""
    return subject, body

# ----------------------------
# SEND EMAIL
# ----------------------------
def send_email(recipient_email, role):
    company = extract_company(recipient_email)
    subject, body = generate_email(company, role)

    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.set_content(body)

    # Attach resume
    with open(RESUME_PATH, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(RESUME_PATH)
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)

    print(f"✅ Email sent to {recipient_email} ({company})")

# ----------------------------
# RUN
# ----------------------------
if __name__ == "__main__":
    recipient = "careers-mbfsindia@mercedes-benz.com"
    role = "Business Intelligence Analyst"

    send_email(recipient, role)
