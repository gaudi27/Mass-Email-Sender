#imports
import gspread
import os
import sys
import json
import re
import mimetypes
import google.auth.exceptions
from googleapiclient.errors import HttpError
from tkinter import messagebox
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
from tkinter import messagebox

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
    creds = None
    # Token.json stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPE)
    # If there are no valid credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  # Refresh the token
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPE)
            creds = flow.run_local_server(port=0, access_type='offline')

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds

'''def get_credentials():
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
    return creds'''


def read_sheet(service, spreadsheet_id, range_name):
    """
    Read data from a Google Sheet in a specific range.
    We are now reading all data from the first column (A), stopping after a certain number of consecutive empty cells.
    """
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get("values", [])
    
    email_list = []
    empty_cell_count = 0
    empty_threshold = 50  # Stop after encountering 50 consecutive empty cells

    for row in values:
        if row and row[0]:  # If the row is non-empty and has a valid email in the first column
            email_list.append(row[0])
            empty_cell_count = 0  # Reset empty cell counter
        else:
            empty_cell_count += 1
            if empty_cell_count >= empty_threshold:
                break  # Stop reading further after too many consecutive empty cells

    return email_list

def get_email_list(sheet_url):
    """Get email list from Google Sheet."""
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    sheet_id = extract_sheet_id(sheet_url)
    try:
        email_list = read_sheet(service, sheet_id, "A:A")
        print("Retrieved email list:", email_list)
        return email_list
    
    except HttpError as error:
        if error.resp.status in [403, 404]:
            messagebox.showerror("Access Error", "Cannot access the Google Sheet. Please set the share permissions to 'Anyone with link can edit'.")
        else:
            messagebox.showerror("Error", f"An error occurred: {error}")
        return []

    except google.auth.exceptions.GoogleAuthError as e:
        messagebox.showerror("Authentication Error", f"Authentication failed: {str(e)}")
        return []

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
        return []


def convert_to_html(text):
    paragraphs = text.split("\n\n")
    html_paragraphs = []

    for paragraph in paragraphs:
        lines = paragraph.split("\n")
        html_paragraphs.append("<p>" + "<br>".join(lines) + "</p>")

    return "".join(html_paragraphs)


def send_email(recipient_email, subject, body):
    from_email = 'updates@viaka.net'
    from_password = 'hzmh ebok orna konf'

    # Convert plain text with newlines to HTML
    body_html = convert_to_html(body)
    
    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = from_email
    message['To'] = recipient_email
    message['Subject'] = subject
    message.attach(MIMEText(body_html, 'html'))
    try:
        # Connect to the server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)

        # Send the email
        server.send_message(message)
        server.quit()
    except smtplib.SMTPRecipientsRefused as e:
        print(f"Error sending email to {recipient_email}: {e}")
    except Exception as e:
        print(f"An unexpected error occurred while sending to {recipient_email}: {e}")

def is_valid_email(email):
    """
    Validate the email format using a regular expression.
    """
    email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
    return re.match(email_regex, email)

def send_emails_in_batches(sheet_url, subject, body, batch_size=20, delay=3600, update_ui=None):
    """Send emails in batches."""
    email_list = get_email_list(sheet_url)
    total_emails = len(email_list)
    
    for i in range(0, total_emails, batch_size):
        batch = email_list[i:i + batch_size]
        batch_successful = 0

        batch = email_list[i:i + batch_size]
        for email in batch:
            if not is_valid_email(email):
                print("skipping email")
                batch_successful += 1
                continue
            try:
                send_email(email, subject, body)
            except Exception as e:
                    print(f"Skipping email {email} due to error: {e}")

        update_ui(f'Sent batch {i // batch_size + 1}')
        
        if i + batch_size < total_emails:
            update_ui(f"Waiting {delay} seconds before sending the next batch...")
            time.sleep(delay)

        print(f'Sent batch {i // batch_size + 1}')
    messagebox.showinfo("Emails Sent", "All emails have been sent successfully!")
