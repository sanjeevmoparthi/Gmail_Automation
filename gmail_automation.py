"""
gmail_automation.py
-------------------
Monthly Gmail Automation using Python + Gmail API
Author: Sanjeev
"""

import os
import base64
import logging
import pandas as pd
from datetime import datetime

from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# ==============================
# CONFIGURATION
# ==============================

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'
EXCEL_FILE = 'email_list.xlsx'
LOG_FILE = 'email_automation.log'

EMAIL_SUBJECT = "Monthly Update"


# ==============================
# LOGGING SETUP
# ==============================

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)


# ==============================
# GMAIL AUTHENTICATION
# ==============================

def gmail_authenticate():
    """
    Authenticate user and return Gmail service
    """
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)


# ==============================
# SEND EMAIL FUNCTION
# ==============================

def send_email(service, to_email, subject, body):
    """
    Send email using Gmail API
    """
    message = MIMEText(body)
    message['to'] = to_email
    message['subject'] = subject

    raw_message = base64.urlsafe_b64encode(
        message.as_bytes()
    ).decode()

    service.users().messages().send(
        userId='me',
        body={'raw': raw_message}
    ).execute()


# ==============================
# MAIN AUTOMATION LOGIC
# ==============================

# def main():
#     logging.info("===== Gmail Monthly Automation Started =====")

#     try:
#         # Authenticate Gmail
#         service = gmail_authenticate()
#         logging.info("Gmail authentication successful")

#         # Load Excel data
#         df = pd.read_excel(EXCEL_FILE)
#         logging.info(f"Loaded {len(df)} records from Excel")

#         # Loop through recipients
#         for index, row in df.iterrows():
#             email = str(row['email']).strip()
#             name = str(row['name']).strip()

#             email_body = f"""
# Hi {name},

# This is your monthly automated update.

# Date: {datetime.today().strftime('%d-%m-%Y')}

# Thank you,
# Sanjeev
# """

#             try:
#                 send_email(service, email, EMAIL_SUBJECT, email_body)
#                 logging.info(f"Email sent successfully to {email}")

#             except Exception as e:
#                 logging.error(f"Failed to send email to {email} | Error: {str(e)}")

#     except Exception as main_error:
#         logging.critical(f"Automation failed | Error: {str(main_error)}")

#     logging.info("===== Gmail Monthly Automation Completed =====")


# # ==============================
# # SCRIPT ENTRY POINT
# # ==============================

# if __name__ == "__main__":
#     main()
def generate_html_report(report_data):
    html = """
    <html>
    <head>
        <title>Email Automation Report</title>
        <style>
            body { font-family: Arial; padding: 20px; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .sent { color: green; font-weight: bold; }
            .failed { color: red; font-weight: bold; }
        </style>
    </head>
    <body>
        <h2>📧 Gmail Automation Report</h2>
        <p>Generated on: """ + datetime.now().strftime("%d-%m-%Y %H:%M:%S") + """</p>

        <table>
            <tr>
                <th>Email</th>
                <th>Name</th>
                <th>Status</th>
                <th>Sent Time</th>
            </tr>
    """

    for row in report_data:
        status_class = "sent" if row["status"] == "Sent" else "failed"
        html += f"""
        <tr>
            <td>{row['email']}</td>
            <td>{row['name']}</td>
            <td class="{status_class}">{row['status']}</td>
            <td>{row['time']}</td>
        </tr>
        """

    html += """
        </table>
    </body>
    </html>
    """

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
def main():
    logging.info("===== Gmail Monthly Automation Started =====")

    report_data = []   # 👈 store report info

    try:
        service = gmail_authenticate()
        logging.info("Gmail authentication successful")

        df = pd.read_excel(EXCEL_FILE)
        logging.info(f"Loaded {len(df)} records from Excel")

        for _, row in df.iterrows():
            email = str(row['email']).strip()
            name = str(row['name']).strip()
            sent_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            email_body = f"""
Hi {name},

This is your monthly automated update.

Date: {datetime.today().strftime('%d-%m-%Y')}

Thank you,
Sanjeev
"""

            try:
                send_email(service, email, EMAIL_SUBJECT, email_body)
                logging.info(f"Email sent successfully to {email}")

                report_data.append({
                    "email": email,
                    "name": name,
                    "status": "Sent",
                    "time": sent_time
                })

            except Exception as e:
                logging.error(f"Failed to send email to {email} | Error: {str(e)}")

                report_data.append({
                    "email": email,
                    "name": name,
                    "status": "Failed",
                    "time": sent_time
                })

    except Exception as main_error:
        logging.critical(f"Automation failed | Error: {str(main_error)}")

    # ✅ Generate HTML verification report
    generate_html_report(report_data)

    logging.info("===== Gmail Monthly Automation Completed =====")

if __name__ == "__main__":
    main()