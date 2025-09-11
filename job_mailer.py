import os
import time
import base64
import random
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    creds = None
    # ✅ Reuse token.json (no login prompt every time)
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # Refresh or request new login only if needed
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

# Create email with attachment
def create_message_with_attachment(sender, to, subject, body_text, file_path):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    # Attach text body
    message.attach(MIMEText(body_text, "plain"))

    # Attach resume
    if file_path and os.path.exists(file_path):
        with open(file_path, "rb") as f:
            mime = MIMEBase("application", "octet-stream")
            mime.set_payload(f.read())
        encoders.encode_base64(mime)
        mime.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
        message.attach(mime)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}

# Send email
def send_email(service, user_id, message, email):
    try:
        service.users().messages().send(userId=user_id, body=message).execute()
        return True
    except Exception as e:
        error_text = str(e)
        print(f"❌ Failed to send to {email}: {error_text}")

        # ✅ Log bounced/invalid emails
        with open("bounced_emails.txt", "a") as bfile:
            bfile.write(f"{email} | {error_text}\n")

        return False

# Extract name from email
def extract_name(email):
    local_part = email.split("@")[0]
    parts = local_part.replace("_", ".").split(".")
    if parts:
        name = parts[0].capitalize()
        if name.isalpha():
            return name
    return None

def main():
    service = get_gmail_service()
    sender = "your_email@gmail.com"

    # ✅ Subjects about Cloud Fresher
    subjects = [
        "Application for Cloud Fresher Role",
        "Cloud Fresher Job Application",
        "Seeking Opportunities as Cloud Fresher",
        "Entry-Level Cloud Computing Job Inquiry",
        "Cloud Fresher Position – Application",
        "Fresher Cloud Role | Job Application",
    ]

    # ✅ Body templates
    bodies = [
        """Hello {name},

I am writing to express my interest in Cloud Computing opportunities at your organization. 
As a fresher with strong fundamentals in cloud and related technologies, I am eager to contribute and grow.

Looking forward to your guidance.

Best regards,  
Nitish
""",
        """Dear {name},

I hope you are doing well. I am seeking entry-level opportunities in Cloud Computing. 
With my educational background and training, I am confident in my ability to learn quickly and add value.

Sincerely,  
Nitish
""",
        """Hi {name},

This is Nitish, reaching out regarding potential openings for fresher roles in Cloud Computing. 
I am enthusiastic about building my career in this domain and would be glad to connect.

Thanks & Regards,  
Nitish
"""
    ]

    # 📂 Resume path
    resume_path = "resume.pdf"

    # Read HR emails
    with open("HRMail.txt", "r") as f:
        emails = [line.strip().strip(",") for line in f if line.strip()]

    # Already sent emails
    sent_emails = set()
    if os.path.exists("sent_log.txt"):
        with open("sent_log.txt", "r") as f:
            sent_emails = set(line.strip() for line in f)

    # ✅ Do NOT resend to sent emails
    unsent_emails = [e for e in emails if e not in sent_emails]

    # Daily limit
    daily_limit = 20
    today_batch = unsent_emails[:daily_limit]

    print(f"📩 Sending {len(today_batch)} emails today...")

    count = 0
    for email in today_batch:
        if count >= daily_limit:
            print("⏹️ Daily limit reached. Stopping for today.")
            break

        subject = random.choice(subjects)
        body_template = random.choice(bodies)
        name = extract_name(email) or "there"
        body = body_template.format(name=name)

        msg = create_message_with_attachment(sender, email, subject, body, resume_path)
        if send_email(service, "me", msg, email):
            print(f"✅ Sent to {email} | Subject: {subject}")
            with open("sent_log.txt", "a") as log:
                log.write(email + "\n")
        else:
            print(f"⚠️ Skipped {email}")

        count += 1
        time.sleep(random.randint(60, 180))

    print("🎉 Finished today's batch.")

if __name__ == '__main__':
    main()
