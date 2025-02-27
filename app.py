import os
import zipfile
import shutil
import email
from email import policy
from email.parser import BytesParser
import logging
import streamlit as st
# from volcenginesdkarkruntime import Ark
from openai import OpenAI
import openai

# Set API Key for authentication
os.environ["ARK_API_KEY"] = "5ade76c2-9629-4076-aebd-3550719382e6"  # Replace with actual API Key

# Initialize Volcano Ark API client
client = OpenAI(
    api_key = os.environ.get("ARK_API_KEY"),
    base_url = "https://ark.cn-beijing.volces.com/api/v3",
)

# openai.api_key = os.getenv("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")

# client = OpenAI(
#     # defaults to os.environ.get("OPENAI_API_KEY")
#     api_key="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",
# )


# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


# Function to extract ZIP file
def extract_zip(zip_path, extract_to="temp_emails"):
    """Extracts all .eml files from a ZIP archive."""
    os.makedirs(extract_to, exist_ok=True)
    
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall(extract_to)

    return [os.path.join(extract_to, file) for file in os.listdir(extract_to) if file.endswith(".eml")]


# Function to extract email content
def extract_email_content(eml_file):
    """Extracts email subject, sender, and body from a .eml file, handling missing body cases."""
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
                    body = part.get_payload(decode=True).decode(part.get_content_charset(), errors="ignore")
                    break  

        if body is None:
            if msg.get_body(preferencelist=("plain")):
                body = msg.get_body(preferencelist=("plain")).get_content()
            elif msg.get_body(preferencelist=("html")):
                body = "HTML Email: " + msg.get_body(preferencelist=("html")).get_content()
            else:
                body = "No readable content available."

        logging.info(f"Extracted Email: {subject} from {sender}")

        return {"subject": subject, "sender": sender, "body": body}

    except Exception as e:
        logging.error(f"Error processing {eml_file}: {e}")
        return {"subject": "Error", "sender": "Error", "body": "Could not process email."}


# Function to analyze email with VolcEngine AI
# Function to analyze emails in one batch
def analyze_emails_with_volcengine(all_emails, user_prompt):
    """Sends all email body texts as a single batch to VolcEngine AI for one combined summary."""
    try:
        # Combine all emails into a single input text
        combined_email_text = "\n\n---\n\n".join(all_emails)  # Separate emails with a delimiter

        # save combined emails as a text file
        with open("combined_emails.txt", "w") as file:
            file.write(combined_email_text)

        # Send one request instead of multiple
        completion = client.chat.completions.create(
            model="ep-20250217174902-6shq5",
            messages=[
                {"role": "system", "content": user_prompt},
                {"role": "user", "content": combined_email_text},
            ],
        )

        return completion.choices[0].message.content

    except Exception as e:
        return f"Error: {str(e)}"


# ----------------- Streamlit UI -----------------
st.set_page_config(page_title="📩 AI Email Analyzer", layout="wide")


# Custom CSS for DeepSeek-like UI
# Custom CSS for Full Dark Mode (Fix White Text Issues)
st.markdown("""
    <style>
        /* Full dark mode background */
        .stApp {
            background-color: #121212;
            color: #ffffff;
        }

        /* Title */
        h1, h2, h3, h4, h5, h6 {
            color: #ffffff;
            font-weight: bold;
            text-align: center;
        }

        /* Instructions and normal text */
        p, span, div {
            color: #dddddd !important;
        }

        /* Text input fields */
        .stTextInput > div > div > input, 
        .stTextArea > div > textarea {
            background-color: #1e1e1e;
            color: #ffffff !important;
            border: 1px solid #555555;
            border-radius: 8px;
        }

        /* Buttons */
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

        /* Uploaded file name */
        .uploadedFileName {
            font-size: 16px;
            font-weight: bold;
            color: #ffffff;
            background-color: #333333;
            padding: 5px 10px;
            border-radius: 8px;
            display: inline-block;
            margin-top: 10px;
        }

        /* Processing message */
        .loading {
            text-align: center;
            font-size: 18px;
            font-weight: bold;
            color: #ffffff;
            background-color: #333333;
            padding: 10px;
            border-radius: 8px;
        }

        /* Success messages */
        .success {
            background-color: #1b5e20;
            color: #ffffff;
            font-weight: bold;
            padding: 10px;
            border-radius: 8px;
            text-align: center;
        }

        /* Error messages */
        .error {
            background-color: #b71c1c;
            color: #ffffff;
            font-weight: bold;
            padding: 10px;
            border-radius: 8px;
            text-align: center;
        }

        /* Expander for AI analysis results */
        .st-expander {
            background-color: #222222 !important;
            color: #ffffff !important;
            border: 1px solid #444444 !important;
            border-radius: 8px;
        }

        /* Expander title */
        .st-expander > summary {
            color: #ffffff !important;
        }

        /* Scrollbars */
        ::-webkit-scrollbar {
            width: 8px;
        }
        ::-webkit-scrollbar-track {
            background: #121212;
        }
        ::-webkit-scrollbar-thumb {
            background: #333333;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #555555;
        }
    </style>
""", unsafe_allow_html=True)


# UI Header
st.title("📩 AI-Powered Email Analyzer")
st.write("Analyze your emails instantly using VolcEngine AI. Upload a ZIP file and get insights!")


# User Input: Custom AI Prompt
user_prompt = st.text_area(
    "🔍 Enter your analysis prompt",
    "首先，请说明这些邮件涉及了哪些项目。接着，请判定是否存在需要管理层关注的系统性风险。如果有，请说明管理层该找谁进一步详细了解风险，请给出管理层需要联系的人员名单并说明人员对应职责。",
    height=100
)


# File Upload
uploaded_file = st.file_uploader("📂 Upload a ZIP file containing .eml emails", type=["zip"])


if uploaded_file:
    with st.spinner("Extracting emails..."):
        eml_files = extract_zip(uploaded_file)

    if not eml_files:
        st.error("❌ No .eml files found in the ZIP file.")
    else:
        st.success(f"✅ Extracted {len(eml_files)} emails. Processing...")

        all_email_texts = []
        for eml_file in eml_files:
            email_data = extract_email_content(eml_file)
            all_email_texts.append(email_data["body"])
        
        # Send all emails together for one combined summary
        combined_analysis = analyze_emails_with_volcengine(all_email_texts, user_prompt)
        
        # Display Single Summary
        st.subheader("📊 AI Analysis Summary")
        st.markdown(f"**🤖 AI Summary:**\n\n{combined_analysis}")

        
        # Display Results
        # st.subheader("📊 AI Analysis Results")
        # for result in results:
        #     with st.expander(f"📩 {result['subject']} ({result['sender']})"):
        #         st.write(f"**📨 Sender:** {result['sender']}")
        #         st.markdown(
        #             f"""
        #             <div style="color: #222; font-size: 16px; line-height: 1.6;">
        #                 <b>🤖 AI Analysis:</b><br>
        #                 {result['analysis']}
        #             </div>
        #             """,
        #             unsafe_allow_html=True
        #         )

                # st.write(f"**🤖 AI Analysis:** {result['analysis']}")

        # Clean up extracted files
        # shutil.rmtree(os.path.dirname(eml_files[0]))
        # Safely clean up the extracted email folder
        extract_folder = "temp_emails"
        if os.path.exists(extract_folder) and os.path.isdir(extract_folder):
            shutil.rmtree(extract_folder)

