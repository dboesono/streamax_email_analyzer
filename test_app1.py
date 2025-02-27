import os
import email
import poplib
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from datetime import datetime
import logging
import streamlit as st
from openai import OpenAI
import openai

# -------------------- VolcEngine API Setup --------------------
from config import ARK_API_KEY

os.environ["ARK_API_KEY"] = ARK_API_KEY  
client = OpenAI(
    api_key=os.environ.get("ARK_API_KEY"),
    base_url="https://ark.cn-beijing.volces.com/api/v3",
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# -------------------- Email Analysis Functions --------------------
def fetch_emails_from_coremail(email_user, email_pass, start_date, end_date):
    EMAIL_HOST = "mail.streamax.com"  # Coremail POP3 server
    EMAIL_PORT = 995  # POP3 SSL port
    mail = poplib.POP3_SSL(EMAIL_HOST, EMAIL_PORT)
    mail.user(email_user)
    mail.pass_(email_pass)

    total_messages = len(mail.list()[1])
    emails = []

    for i in range(total_messages, 0, -1):
        raw_email_bytes = b"\n".join(mail.retr(i)[1])  # Fetch raw email
        msg = BytesParser(policy=policy.default).parsebytes(raw_email_bytes)  # Parse email

        # Parse email date and convert to datetime
        email_date_str = msg["date"]
        try:
            email_date_parsed = parsedate_to_datetime(email_date_str)
            email_date = email_date_parsed.replace(tzinfo=None)
        except Exception as e:
            continue

        # Compare only the date portion
        if email_date.date() < start_date or email_date.date() > end_date:
            continue

        # Extract email content (subject, sender, body)
        subject = msg["subject"] or "No Subject"
        sender = msg["from"]
        body = None
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    body = part.get_payload(decode=True).decode(part.get_content_charset(), errors="ignore")
                    break
        if not body:
            body = "No readable content available."

        emails.append({"subject": subject, "sender": sender, "body": body})

    mail.quit()
    return emails

def analyze_emails_with_volcengine(all_emails, user_prompt):
    try:
        # Combine all emails into one text to send to VolcEngine
        combined_email_text = "\n\n---\n\n".join([email["body"] for email in all_emails])

        # Send to VolcEngine API for analysis (example of sending a request)
        response = client.chat.completions.create(
            model="ep-20250217174902-6shq5",  # Replace with your model ID
            messages=[
                {"role": "system", "content": user_prompt},
                {"role": "user", "content": combined_email_text},
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

# -------------------- Streamlit UI --------------------
st.set_page_config(page_title="📩 Email Analyzer", layout="wide")
st.title("📩 Automated Email Analyzer")

# User inputs email and timeframe
email_user = st.text_input("Coremail Email Address")
email_pass = st.text_input("Coremail Password", type="password")

# Timeframe for email filtering
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date", value=datetime(2025, 2, 19).date())
with col2:
    end_date = st.date_input("End Date", value=datetime(2025, 2, 26).date())

# User prompt for analysis
user_prompt = st.text_area("Enter your analysis prompt", "Analyze the emails and provide a summary.", height=100)

if st.button("Analyze Emails"):
    if email_user and email_pass:
        # Fetch emails directly from Coremail
        emails = fetch_emails_from_coremail(email_user, email_pass, start_date, end_date)
        if emails:
            # Analyze the emails
            analysis_result = analyze_emails_with_volcengine(emails, user_prompt)
            st.subheader("Analysis Result")
            st.write(analysis_result)
        else:
            st.error("No emails found in the specified timeframe.")
    else:
        st.error("Please enter your email and password.")


