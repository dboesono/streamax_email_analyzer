import poplib
import email
import os
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from datetime import datetime

# Email account credentials
EMAIL_HOST = "mail.streamax.com"  # POP3 server
EMAIL_PORT = 995  # SSL POP3 port
EMAIL_USER = "Darren@streamax.com"  # Your email address
EMAIL_PASS = "Luojielun28db_coremail"  # Your email password

# Directories for saving emails and attachments
EMAIL_DIR = "emails"
ATTACHMENT_DIR = "attachments"
os.makedirs(EMAIL_DIR, exist_ok=True)
os.makedirs(ATTACHMENT_DIR, exist_ok=True)

# Define the timeframe for filtering emails (adjust these as needed)
# For example, only process emails from Jan 1, 2023 to Dec 31, 2023
START_DATE = datetime(2025, 2, 19)
END_DATE = datetime(2025, 2, 26+1)

def connect_to_pop3():
    """Connect to the POP3 email server using SSL."""
    try:
        mail = poplib.POP3_SSL(EMAIL_HOST, EMAIL_PORT)
        mail.user(EMAIL_USER)
        mail.pass_(EMAIL_PASS)
        print(f"Connected to {EMAIL_HOST}, Total emails: {len(mail.list()[1])}")
        return mail
    except Exception as e:
        print("Failed to connect to POP3 server:", e)
        return None

def fetch_and_save_emails(mail, num_messages=10):
    """Fetch and save emails within a specific timeframe as .eml files."""
    total_messages = len(mail.list()[1])
    num_messages = min(num_messages, total_messages)  # Avoid out-of-range issues

    print(f"Fetching the last {num_messages} emails...\n")

    for i in range(total_messages, total_messages - num_messages, -1):
        raw_email_bytes = b"\n".join(mail.retr(i)[1])  # Fetch raw email
        msg = BytesParser(policy=policy.default).parsebytes(raw_email_bytes)  # Parse email

        # Extract metadata
        subject = msg["subject"] or "No Subject"
        sender = msg["from"]
        email_date_str = msg["date"]

        # Parse the email date and convert to a naive datetime for comparison
        try:
            email_date_parsed = parsedate_to_datetime(email_date_str)
            email_date = email_date_parsed.replace(tzinfo=None)
        except Exception as e:
            print(f"Skipping Email {i} due to date parsing error: {email_date_str}")
            continue

        # Check if the email date is within the desired timeframe
        if email_date < START_DATE or email_date > END_DATE:
            print(f"Skipping Email {i}: {subject} from {sender} on {email_date_str}")
            continue

        print(f"Saving Email {i}: {subject} from {sender} on {email_date_str}")

        # Save full email as .eml file
        eml_filename = os.path.join(EMAIL_DIR, f"email_{i}.eml")
        with open(eml_filename, "wb") as f:
            f.write(raw_email_bytes)
        print(f"Saved: {eml_filename}")

        # Save attachments (if any)
        save_attachments(msg)

    print("\nEmail fetching complete!")

def save_attachments(msg):
    """Saves email attachments to a local folder."""
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue

        filename = part.get_filename()
        if filename:
            filepath = os.path.join(ATTACHMENT_DIR, filename)
            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            print(f"Saved attachment: {filepath}")

if __name__ == "__main__":
    mail = connect_to_pop3()
    if mail:
        # Fetch all emails, but only those within the specified timeframe will be saved
        fetch_and_save_emails(mail, num_messages=len(mail.list()[1]))
        mail.quit()



