#imports
import gspread
import os
import sys
import json
import re
import mimetypes
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, MediaFileUpload
from googleapiclient.errors import HttpError
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

def extract_sheet_id(url):
    """
    Extracts the Google Sheets ID from the provided URL.
    
    Args:
        url (str): The URL of the Google Sheet.
        
    Returns:
        str: The extracted Google Sheets ID.
    """
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid Google Sheets URL")


SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]

Sheet_URL = "https://docs.google.com/spreadsheets/d/1jLXfWn1WZBZk7yH0BPQQ-50SEvnCEPLj_iGZS1fwaxI/edit?pli=1&gid=0#gid=0"

def get_credentials():
    """Get user credentials for Google Sheets API."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPE)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPE)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds


def read_sheet(service, spreadsheet_id, range_name):
    """Read data from a Google Sheet."""
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get("values", [])

def get_email_list(sheet_url):
    """Get email list from Google Sheet."""
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    sheet_id = extract_sheet_id(sheet_url)
    email_list = read_sheet(service, sheet_id, "A1:A1000")
    return [email[0] for email in email_list]


def send_email(recipient_email, subject, body):
    from_email = 'your@email.com'
    from_password = 'yourpassword'

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = from_email
    message['To'] = recipient_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    # Connect to the server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, from_password)

    # Send the email
    server.send_message(message)
    server.quit()

def send_emails_in_batches(sheet_url, subject, body, batch_size=20, delay=3600):
    """Send emails in batches."""
    email_list = get_email_list(sheet_url)
    total_emails = len(email_list)
    
    for i in range(0, total_emails, batch_size):
        batch = email_list[i:i + batch_size]
        for email in batch:
            send_email(email, subject, body)
        print(f'Sent batch {i // batch_size + 1}')
        time.sleep(delay)
