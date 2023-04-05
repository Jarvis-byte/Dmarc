import os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.cloud import logging

def check_gmail_for_new_emails():
    # Authenticate with the Gmail API using OAuth 2.0 credentials
    credentials = Credentials.from_authorized_user_file(os.environ['GOOGLE_APPLICATION_CREDENTIALS'], ['https://www.googleapis.com/auth/gmail.readonly'])
    service = build('gmail', 'v1', credentials=credentials)

    # Retrieve the list of messages in your inbox
    messages = service.users().messages().list(userId='me', q='newer_than:5m').execute().get('messages', [])

    # If there are new messages, log them to Cloud Logging
    if messages:
        logging_client = logging.Client()
        logger = logging_client.logger('my-logger')
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            subject = next((header['value'] for header in msg['payload']['headers'] if header['name'] == 'Subject'), '')
            sender = next((header['value'] for header in msg['payload']['headers'] if header['name'] == 'From'), '')
            date = next((header['value'] for header in msg['payload']['headers'] if header['name'] == 'Date'), '')
            logger.log_text(f'New email: subject={subject}, sender={sender}, date={date}')

# This function will be triggered by the scheduled job
def run_check_email_job(request):
    check_gmail_for_new_emails()
    return 'OK'