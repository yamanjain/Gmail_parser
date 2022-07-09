# from __future__ import print_function

from bs4 import BeautifulSoup
import base64
import json
import math
import os.path
import sys
import re

from io import BytesIO

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
from pandas import read_excel
import fitz

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    cwd = os.getcwd()
    try:
        if os.path.exists(cwd + os.path.sep + 'token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(cwd + os.path.sep + 'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(cwd + os.path.sep + 'token.json', 'w') as token:
                token.write(creds.to_json())

    except Exception as error:
        print(f'An error occurred during authentication: {error}')

    try:
        service = build('gmail', 'v1', credentials=creds, static_discovery=False)
        all_emails = fetch_emails(service)
        output_list = []
        for result in all_emails:
            # Fetch Email object with raw email body and attachment ids
            msg = service.users().messages().get(userId="me", id=result['id'], format="full").execute()
            payload = msg['payload']
            parts = payload.get('parts')
            if parts is None:
                # Simpler method without recursion
                email_body = parse_msg_body(msg)
            else:
                # Get body using recursion
                [body_message, body_html] = processParts(parts)
                email_body_text = str(body_message, 'utf-8')
                email_html = str(body_html, 'utf-8')
                email_html_text = clean_html(email_html)
                email_body = email_body_text + email_html_text

            email_headers = get_headers(msg)
            attachment_obj = parse_attachment_as_dict(service, msg, result['id'])
            bank_ref_no = ['NA']
            if 'BankRefNo' in attachment_obj:
                bank_ref_no = attachment_obj['BankRefNo']
            # if 'Child Claim ref number' in attachment_obj and math.isnan(attachment_obj['Child Claim ref number'][0]):
            #     attachment_obj['Child Claim ref number'][0] = "NA"
            # change all NaN to "NA" in the attachment_obj dictionary
            if 'Child Claim ref number' in attachment_obj:
                stringify = json.dumps(attachment_obj)
                regex = re.compile(r'\bnan\b', flags=re.IGNORECASE)
                stringify = re.sub(regex, r'"NA"', stringify)
                attachment_obj = json.loads(stringify)

            # Create a python data structure with all the info and convert it to json and write to a file
            output_list.append({'bank_ref_no': bank_ref_no, 'email_headers': email_headers, 'email_body': email_body,
                                'attachment_obj': attachment_obj})
        with open("payload.json", "w") as outfile:
            json.dump(output_list, outfile)

    except HttpError as error:
        print(f'An error occurred: {error}')


def get_headers(msg):
    headers_needed = ['From', 'To', 'Date', 'Subject']
    res_dict = {}
    headers_list = msg.get("payload").get('headers')
    for header in headers_list:
        if header['name'] in headers_needed:
            res_dict[header['name']] = header['value']

    return res_dict


def fetch_emails(service):
    """
    Get all emails based on a filter passed in args
    """
    try:
        # Filter q string passed from caller.
        args = sys.argv[1:]
        # Call the Gmail API
        # results = service.users().messages().list(userId='me', q="from:NIAHO@newindia.co.in niaho newer_than:4d").execute()
        # results = service.users().messages().list(userId='me', q='from:support@icicilombard.com subject:fund transfer for motor claim newer_than:8d').execute()
        results = service.users().messages().list(userId='me', q="niaho tds newer_than:15d").execute()
        # results = service.users().messages().list(userId='me', q=args[0]).execute()
        if 'messages' in results:
            return results['messages'] or []
        else:
            return []
    except HttpError as error:
        print(f'An error occurred: {error}')


def parse_msg_body(msg):
    """
    Find the email body from the respective email object
    Navigate to payload > body > data, if not present it should be available as a part object
    Navigate to the part[0] to get the available parts and retrieve the email body
    Navigate to the part[0] -> part[0] to get the available parts and retrieve the email body
    If the email body is not available, from above three then return the snippet.
    """
    try:
        if msg.get("payload").get("body").get("data"):
            return base64.urlsafe_b64decode(msg.get("payload").get("body").get("data").encode("ASCII")).decode("utf-8")
        elif msg.get("payload").get("parts")[0].get("body").get("data"):
            return base64.urlsafe_b64decode(
                msg.get("payload").get("parts")[0].get("body").get("data").encode("ASCII")).decode("utf-8")
        elif msg.get("payload").get("parts")[0].get("parts")[0].get("body").get("data"):
            return base64.urlsafe_b64decode(
                msg.get("payload").get("parts")[0].get("parts")[0].get("body").get("data")).decode("utf-8")
        return msg.get("snippet")
    except Exception as e:
        err_msg = f"Error occurred while fetching email body: {e}"
        print(err_msg)
        return err_msg


def parse_attachment_as_dict(service, msg, msg_id):
    """
    Parse attachment as a dict.
    """
    # Check all attachments and find the right one (MS Excel)
    parts = msg.get('payload').get('parts')
    if not parts:
        return {}

    all_parts = []
    for p in parts:
        if p.get('parts'):
            all_parts.extend(p.get('parts'))
        else:
            all_parts.append(p)

    att_parts = [p for p in all_parts if
                 (p['mimeType'] == 'application/vnd.ms-excel' or p['mimeType'] == 'application/octet-stream' or p[
                     'mimeType'] == 'application/pdf')]
    # filenames = [p['filename'] for p in att_parts]
    attachment_obj = {}

    # Find the correct attachment using attachment id
    for part in att_parts:
        data = part['body'].get('data')
        attachment_id = part['body'].get('attachmentId')
        if not data:
            # Retrieve the attachment separately since it's not part of the message object
            att = service.users().messages().attachments().get(
                userId='me', id=attachment_id, messageId=msg_id).execute()
            data = att['data']
            str_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            # Convert string data to dataframe
            if p['mimeType'] == 'application/pdf':
                with fitz.open(stream=str_data, filetype="pdf") as doc:
                    text = ""
                    for page in doc:
                        text += page.get_text()
                attachment_obj = {}
                attachment_obj['pdf_text'] = text
            else:
                try:
                    df = read_excel(BytesIO(str_data))
                    attachment_obj = {}
                    # Create dict with key as Column name and value as the actual value
                    for col in df.columns:
                        val = list(df[col])
                        attachment_obj[col] = val
                    # if df.len() > 2:
                    #     # df.groupby('BankRefNo').apply(lambda x: x.to_dict(orient='r')).to_dict()
                    #     attachment_obj = df.todict(orient='records')
                    break
                except Exception as e:
                    attachment_obj['Error'] = str(e)

    return attachment_obj


def processParts(parts):
    if not 'body_message' in locals():
        body_message = bytearray()
    if not 'body_html' in locals():
        body_html = bytearray()
    for part in parts:
        body = part.get("body")
        data = body.get("data")
        mimeType = part.get("mimeType")
        if mimeType == 'multipart/alternative':
            subparts = part.get('parts')
            [body_message, body_html] = processParts(subparts)
        elif mimeType == 'text/plain' and not isinstance(data, type(None)):
            body_message = base64.urlsafe_b64decode(data)
        elif mimeType == 'text/html' and not isinstance(data, type(None)):
            body_html = base64.urlsafe_b64decode(data)
    return [body_message, body_html]


def clean_html(email_html):
    soup = BeautifulSoup(email_html, features="html.parser")
    # kill all script and style elements
    for script in soup(["script", "style"]):
        script.extract()  # rip it out

    # get text
    text = soup.get_text()

    # break into lines and remove leading and trailing space on each
    lines = (line.strip() for line in text.splitlines())
    # break multi-headlines into a line each
    chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
    # drop blank lines
    text = '\n'.join(chunk for chunk in chunks if chunk)
    return text


def build_json():
    pass


if __name__ == '__main__':
    if os.path.exists("payload.json"):
        os.remove("payload.json")
    main()
