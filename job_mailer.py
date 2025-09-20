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
from googleapiclient.http import BatchHttpRequest
from google.auth.transport.requests import Request

# Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def create_message_with_attachment(sender, to, subject, body_text, file_path):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    message.attach(MIMEText(body_text, "plain"))
    if file_path and os.path.exists(file_path):
        with open(file_path, "rb") as f:
            mime = MIMEBase("application", "octet-stream")
            mime.set_payload(f.read())
        encoders.encode_base64(mime)
        mime.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
        message.attach(mime)
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}

def send_email_callback(request_id, response, exception, email, sent_log_file):
    if exception is not None:
        error_text = str(exception)
        print(f"❌ Failed to send to {email}: {error_text}")
        with open("bounced_emails.txt", "a") as bfile:
            bfile.write(f"{email} | {error_text}\n")
    else:
        print(f"✅ Sent to {email}")
        with open(sent_log_file, "a") as log:
            log.write(email + "\n")

def extract_name(email):
    local_part = email.split("@")[0]
    parts = local_part.replace("_", ".").split(".")
    if parts:
        name = parts[0].capitalize()
        if name.isalpha():
            return name
    return None

def analyze_bounced_emails(bounced_file="bounced_emails.txt"):
    retryable_emails = []
    if not os.path.exists(bounced_file):
        return retryable_emails
    with open(bounced_file, "r") as f:
        for line in f:
            email, error = line.strip().split(" | ", 1)
            if "550" not in error and "does not exist" not in error.lower():
                retryable_emails.append(email)
    return retryable_emails

def main():
    service = get_gmail_service()
    sender = "your_email@gmail.com"
    resume_path = "resume.pdf"
    sent_log_file = "sent_log.txt"
    daily_limit = 50

    # Subjects
    subjects = [
        "Application for Cloud Fresher Role",
        "Cloud Fresher Job Application",
        "Seeking Opportunities as Cloud Fresher",
        "Entry-Level Cloud Computing Job Inquiry",
        "Cloud Fresher Position – Application",
        "Fresher Cloud Role | Job Application",
        "Exploring Cloud Computing Opportunities",
        "Application for Entry-Level Cloud Position",
        "Cloud Fresher – Career Opportunity Inquiry",
        "Interested in Cloud Computing Roles",
        "Cloud Fresher Job Opportunity Application"
    ]

    # Organized body templates as a list of dictionaries
    body_templates = [
        # Formal tone templates
        {
            "tone": "formal",
            "text": """Dear {name},\n\nI am writing to express my interest in entry-level Cloud Computing roles at your organization. As a fresher with a strong foundation in cloud technologies, I am eager to contribute and grow professionally. My resume is attached for your review.\n\nSincerely,\nNitish Jhamb"""
        },
        {
            "tone": "formal",
            "text": """Dear {name},\n\nI am reaching out to inquire about Cloud Computing opportunities for freshers. With my academic background and training in cloud systems, I am confident in my ability to add value to your team. Please find my resume attached.\n\nBest regards,\nNitish Jhamb"""
        },
        {
            "tone": "formal",
            "text": """Dear {name},\n\nI am interested in exploring fresher roles in Cloud Computing at your esteemed organization. My technical skills and enthusiasm for cloud technologies make me a strong candidate. My resume is attached for your consideration.\n\nKind regards,\nNitish Jhamb"""
        },
        {
            "tone": "formal",
            "text": """Dear {name},\n\nI am a fresher with a keen interest in Cloud Computing and am excited to apply for opportunities at your company. My resume, attached, highlights my skills and dedication to this field. I look forward to contributing to your team.\n\nWarm regards,\nNitish Jhamb"""
        },
        # Semi-formal/enthusiastic tone templates
        {
            "tone": "semi-formal",
            "text": """Hello {name},\n\nI’m Nitish, a fresher passionate about Cloud Computing. I’m eager to explore entry-level opportunities at your organization and contribute my skills. Please see my attached resume for more details.\n\nThanks & Regards,\nNitish Jhamb"""
        },
        {
            "tone": "semi-formal",
            "text": """Hi {name},\n\nThis is Nitish, reaching out to explore Cloud Computing roles for freshers. My enthusiasm and foundational knowledge in cloud technologies make me a great fit. My resume is attached for your review.\n\nBest,\nNitish Jhamb"""
        },
        {
            "tone": "semi-formal",
            "text": """Hello {name},\n\nI’m excited to connect regarding fresher opportunities in Cloud Computing. With a solid technical background, I’m eager to contribute to your team’s success. Please find my resume attached.\n\nThank you,\nNitish Jhamb"""
        },
        {
            "tone": "semi-formal",
            "text": """Hi {name},\n\nI’m Nitish, a recent graduate enthusiastic about Cloud Computing. I’d love to explore career opportunities at your organization. My resume, attached, showcases my skills and passion for this field.\n\nBest regards,\nNitish Jhamb"""
        }
    ]

    # Read HR emails
    with open("HRMail.txt", "r") as f:
        emails = [line.strip().strip(",") for line in f if line.strip()]

    # Load sent emails
    sent_emails = set()
    if os.path.exists(sent_log_file):
        with open(sent_log_file, "r") as f:
            sent_emails = set(line.strip() for line in f)

    # Load retryable bounced emails
    retryable_bounced = analyze_bounced_emails()
    unsent_emails = [e for e in emails if e not in sent_emails] + retryable_bounced
    unsent_emails = list(dict.fromkeys(unsent_emails))  # Remove duplicates
    today_batch = unsent_emails[:daily_limit]

    print(f"📩 Sending {len(today_batch)} emails today...")

    batch = service.new_batch_http_request()
    count = 0
    batch_size = 5  # Send 5 emails per batch

    for email in today_batch:
        if count >= daily_limit:
            print("⏹️ Daily limit reached. Stopping for today.")
            break

        subject = random.choice(subjects)
        body_template = random.choice(body_templates)["text"]
        name = extract_name(email) or "there"
        body = body_template.format(name=name)
        msg = create_message_with_attachment(sender, email, subject, body, resume_path)

        batch.add(
            service.users().messages().send(userId="me", body=msg),
            callback=lambda rid, resp, exc, e=email, f=sent_log_file: send_email_callback(rid, resp, exc, e, f)
        )
        count += 1

        if count % batch_size == 0 or count == len(today_batch):
            try:
                batch.execute()
                print(f"📬 Processed batch of {min(batch_size, count % batch_size or batch_size)} emails")
                if count < daily_limit:  # Pause only if more emails remain
                    print("⏸️ Pausing for 1 minute...")
                    time.sleep(60)  # 1-minute pause after every 5 emails
            except Exception as e:
                print(f"❌ Batch failed: {str(e)}")
            batch = service.new_batch_http_request()  # Reset batch

    print("🎉 Finished today's batch.")

if __name__ == '__main__':
    main()