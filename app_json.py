import os
import zipfile
import shutil
import email
import poplib
import json
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from datetime import datetime
import logging
import streamlit as st
from openai import OpenAI
import openai

# -------------------- VolcEngine API Setup --------------------
# Set API Key for authentication (replace with your actual API key)
# from config import ARK_API_KEY

os.environ["ARK_API_KEY"] = "5ade76c2-9629-4076-aebd-3550719382e6"
client = OpenAI(
    api_key=os.environ.get("ARK_API_KEY"),
    base_url="https://ark.cn-beijing.volces.com/api/v3",
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# -------------------- Email Analysis Functions --------------------
def extract_zip(zip_file, extract_to="temp_emails"):
    """Extract all .eml files from a ZIP archive."""
    os.makedirs(extract_to, exist_ok=True)
    with zipfile.ZipFile(zip_file, "r") as zip_ref:
        zip_ref.extractall(extract_to)
    return [os.path.join(extract_to, file) for file in os.listdir(extract_to) if file.endswith(".eml")]

def extract_email_content(eml_file):
    """Extracts subject, sender, and body from a .eml file."""
    try:
        with open(eml_file, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        subject = msg["subject"] if msg["subject"] else "No Subject"
        sender = msg["from"] if msg["from"] else "Unknown Sender"
        body = None
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    charset = part.get_content_charset() or "utf-8"
                    body = part.get_payload(decode=True).decode(charset, errors="ignore")
                    break
        if body is None:
            if msg.get_body(preferencelist=("plain",)):
                body = msg.get_body(preferencelist=("plain",)).get_content()
            elif msg.get_body(preferencelist=("html",)):
                body = "HTML Email: " + msg.get_body(preferencelist=("html",)).get_content()
            else:
                body = "No readable content available."
        logging.info(f"Extracted Email: {subject} from {sender}")
        return {"subject": subject, "sender": sender, "body": body}
    except Exception as e:
        logging.error(f"Error processing {eml_file}: {e}")
        return {"subject": "Error", "sender": "Error", "body": "Could not process email."}

def analyze_emails_with_volcengine(all_emails, user_prompt):
    """
    Converts the list of email JSON objects into a JSON string,
    saves it for debugging, and then sends it to the VolcEngine API
    for analysis based on the provided user prompt.
    """
    combined_emails_json = json.dumps(all_emails, ensure_ascii=False, indent=2)
    # Save the JSON string for debugging purposes
    with open("combined_emails.json", "w", encoding="utf-8") as file:
        file.write(combined_emails_json)
    try:
        completion = client.chat.completions.create(
            model="ep-20250217174902-6shq5",
            messages=[
                {"role": "system", "content": user_prompt},
                {"role": "user", "content": combined_emails_json},
            ],
        )
        return completion.choices[0].message.content, combined_emails_json
    except Exception as e:
        return f"Error: {str(e)}", combined_emails_json

# -------------------- POP3 Email Download Functions --------------------
def connect_to_pop3(email_host, email_port, email_user, email_pass):
    """Connect to the POP3 server using SSL."""
    try:
        mail = poplib.POP3_SSL(email_host, email_port)
        mail.user(email_user)
        mail.pass_(email_pass)
        logging.info(f"Connected to {email_host}, Total emails: {len(mail.list()[1])}")
        return mail
    except Exception as e:
        logging.error("Failed to connect to POP3 server: " + str(e))
        return None

def fetch_and_save_emails(mail, start_date, end_date, output_dir="emails", num_messages=None):
    """
    Fetch emails from the POP3 server, filter by date (user-defined timeframe),
    and save as .eml files.

    The function scans from the most recent email backward. It skips emails newer than end_date,
    processes emails within the timeframe, and once it encounters an email older than start_date,
    it terminates the loop.
    """
    os.makedirs(output_dir, exist_ok=True)
    total_messages = len(mail.list()[1])
    saved_files = []
    scanned = 0

    # Loop from the latest email (highest number) to the oldest (1)
    for i in range(total_messages, 0, -1):
        # Respect num_messages limit if provided (> 0)
        if num_messages and scanned >= num_messages:
            break

        try:
            resp, lines, octets = mail.retr(i)
            raw_email_bytes = b"\n".join(lines)
            msg = BytesParser(policy=policy.default).parsebytes(raw_email_bytes)
            email_date_str = msg["date"]

            try:
                email_date_parsed = parsedate_to_datetime(email_date_str)
                # Convert to a naive datetime (local time) for comparison if tzinfo is present
                if email_date_parsed.tzinfo:
                    email_date = email_date_parsed.astimezone().replace(tzinfo=None)
                else:
                    email_date = email_date_parsed
            except Exception as e:
                logging.error(f"Skipping email {i} due to date parsing error: {email_date_str}")
                continue

            # Skip emails that are newer than the end_date
            if email_date > end_date:
                logging.info(f"Skipping email {i}: email date {email_date} is after end_date {end_date}")
                continue

            # Once we hit an email older than start_date, we can break out of the loop
            if email_date < start_date:
                logging.info(f"Reached email {i} with date {email_date} older than start_date {start_date}. Terminating loop.")
                break

            # Email is within the specified timeframe; process it
            subject = msg["subject"] or "No Subject"
            logging.info(f"Saving email {i}: {subject} on {email_date_str}")
            eml_filename = os.path.join(output_dir, f"email_{i}.eml")
            with open(eml_filename, "wb") as f:
                f.write(raw_email_bytes)
            saved_files.append(eml_filename)
            scanned += 1

        except Exception as e:
            logging.error(f"Error processing email {i}: {e}")

    return saved_files

def create_zip_from_dir(directory, zip_filename="downloaded_emails.zip"):
    """Create a ZIP file from all files in a directory."""
    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for foldername, subfolders, filenames in os.walk(directory):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, directory)
                zipf.write(filepath, arcname)
    return zip_filename

# -------------------- Streamlit UI --------------------
st.set_page_config(page_title="📩 Email Analyzer & Downloader", layout="wide")
st.title("📩 Email Analyzer & Downloader")

# Create two tabs: one for downloading emails and one for analyzing emails
tab1, tab2 = st.tabs(["Download Emails", "Analyze Emails"])

# ---------- Tab 1: Download Emails ----------
with tab1:
    st.header("Download Emails from your POP3 Account")
    with st.form("download_emails_form"):
        email_host = st.text_input("POP3 Server Host", value="mail.streamax.com")
        email_port = st.number_input("POP3 Server Port", value=995, step=1)
        email_user = st.text_input("Email Address")
        email_pass = st.text_input("Email Password", type="password")
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start Date", value=datetime(2025, 2, 19).date())
        with col2:
            end_date = st.date_input("End Date", value=datetime(2025, 2, 26).date())
        num_messages = st.number_input("Number of recent emails to check (0 for all)", min_value=0, value=0, step=1)
        submitted = st.form_submit_button("Download Emails")
    
    if submitted:
        if not email_user or not email_pass:
            st.error("Please provide your email address and password.")
        else:
            # Clear previous emails folder to avoid mixing emails from different queries
            if os.path.exists("emails"):
                shutil.rmtree("emails")
            os.makedirs("emails", exist_ok=True)
            
            st.info("Connecting to POP3 server...")
            mail = connect_to_pop3(email_host, int(email_port), email_user, email_pass)
            if mail:
                st.info("Connected. Fetching emails...")
                # Convert the user-provided dates to datetime objects for comparison
                start_dt = datetime.combine(start_date, datetime.min.time())
                end_dt = datetime.combine(end_date, datetime.max.time())
                saved_files = fetch_and_save_emails(mail, start_dt, end_dt, output_dir="emails", num_messages=num_messages)
                mail.quit()
                if saved_files:
                    st.success(f"Fetched and saved {len(saved_files)} emails.")
                    zip_filename = create_zip_from_dir("emails", zip_filename="downloaded_emails.zip")
                    with open(zip_filename, "rb") as f:
                        st.download_button("Download ZIP file of emails", data=f.read(), file_name=zip_filename, mime="application/zip")
                    # Optionally clear the emails folder after zipping:
                    # shutil.rmtree("emails")
                else:
                    st.error("No emails found in the specified timeframe.")

# ---------- Tab 2: Analyze Emails ----------
with tab2:
    st.header("Analyze Emails with AI")
    st.write("Upload a ZIP file containing .eml emails and enter your analysis prompt below.")
    user_prompt = st.text_area(
        "Enter your analysis prompt",
        "首先，请说明这些邮件涉及了哪些项目。接着，请判定是否存在需要管理层关注的系统性风险。如果有，请说明管理层该找谁进一步详细了解风险，请给出管理层需要联系的人员名单并说明人员对应职责。",
        height=100
    )
    uploaded_file = st.file_uploader("Upload ZIP file containing .eml emails", type=["zip"])
    if uploaded_file:
        with st.spinner("Extracting emails..."):
            # Save the uploaded ZIP file temporarily
            zip_temp_path = "uploaded_emails.zip"
            with open(zip_temp_path, "wb") as f:
                f.write(uploaded_file.read())
            eml_files = extract_zip(zip_temp_path, extract_to="temp_emails")
        if not eml_files:
            st.error("No .eml files found in the ZIP file.")
        else:
            st.success(f"Extracted {len(eml_files)} emails. Processing...")
            # Store each email as a JSON object in a list.
            all_emails = []
            for eml_file in eml_files:
                email_data = extract_email_content(eml_file)
                all_emails.append(email_data)
            # Get the analysis result and the JSON payload used for debugging.
            analysis_result, combined_emails_json = analyze_emails_with_volcengine(all_emails, user_prompt)
            st.subheader("AI Analysis Summary")
            st.markdown(f"**AI Summary:**\n\n{analysis_result}")
            # Provide a download button and a code viewer for the JSON payload.
            st.download_button("Download combined_emails.json", data=combined_emails_json, file_name="combined_emails.json", mime="application/json")
            with st.expander("Show JSON input for debugging"):
                st.code(combined_emails_json, language="json")
            # Clean up temporary files and folders
            if os.path.exists("temp_emails"):
                shutil.rmtree("temp_emails")
            if os.path.exists(zip_temp_path):
                os.remove(zip_temp_path)

# -------------------- Custom CSS Styling (Optional) --------------------
st.markdown("""
    <style>
        .stApp {
            background-color: #121212;
            color: #ffffff;
        }
        body, p, span, div {
            color: #ffffff !important;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #ffffff;
            font-weight: bold;
            text-align: center;
        }
        .stTextInput > div > div > input, 
        .stTextArea > div > textarea {
            background-color: #1e1e1e;
            color: #ffffff !important;
            border: 1px solid #555555;
            border-radius: 8px;
        }
        .stButton>button {
            background-color: #007bff;
            color: white;
            border-radius: 8px;
            padding: 10px 20px;
            font-size: 16px;
        }
        .stButton>button:hover {
            background-color: #0056b3;
        }
    </style>
""", unsafe_allow_html=True)
