import os
import base64
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Scope for sending emails
SCOPES = ['https://www.googleapis.com/auth/gmail.send']


def get_credentials():
    """Get valid credentials from storage or obtain new ones."""
    creds = None
    # Check if token.json exists and load credentials from it
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If no valid credentials, refresh or obtain new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Load credentials from the client secrets file
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save credentials to token.json for future use
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def send_email(to_address, to_name, subject, body):
    """Send an email using the Gmail API."""
    service = build('gmail', 'v1', credentials=get_credentials())

    message = MIMEMultipart()
    message['From'] = ''  # Update with your Gmail address
    message['To'] = to_address
    message['Subject'] = subject

    # Replace the placeholder with the actual name
    personalized_body = body.replace('___', to_name)
    message.attach(MIMEText(personalized_body, 'plain'))

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    send_message = service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
    print(f"Message Id: {send_message['id']}")


def main():
    # Read Excel file
    excel_file_path = ''  # Update with your Excel file path
    df = pd.read_excel(excel_file_path)

    # Email content
    email_subject = 'Test Subject'
    email_body = """Dear ___:

My name is Kunsh Arora. 

Cordially,
Kunsh Arora
"""

    # Send email to each address with personalized greeting
    for index, row in df.iterrows():
        email_address = row['Email']
        to_name = row['Name']
        send_email(email_address, to_name, email_subject, email_body)

    print("Emails sent successfully.")


if __name__ == '__main__':
    main()
