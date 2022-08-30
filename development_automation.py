from __future__ import print_function

import io

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

from oauth2client.service_account import ServiceAccountCredentials
from httplib2 import Http

import pandas as pd
import json

from email.message import EmailMessage
import ssl
import smtplib

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


sender = "aidendawes.spammail@gmail.com"
password = "drgcotrbhurjioyj"
subject = "By-election Alert"


def export_word(real_file_id):
    """Download a Document file in PDF format.
    Args:
        real_file_id : file ID of any workspace document format file
    Returns : IO object with location

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    scopes = ['https://www.googleapis.com/auth/drive.readonly']

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'credentials.json', scopes)

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)

        file_id = real_file_id

        # pylint: disable=maybe-no-member
        request = service.files().export_media(fileId=file_id,
                                               mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(F'Download {int(status.progress() * 100)}.')

    except HttpError as error:
        print(F'An error occurred: {error}')
        file = None

    return file.getvalue()


def export_excel(real_file_id):
    """Download a Document file in PDF format.
    Args:
        real_file_id : file ID of any workspace document format file
    Returns : IO object with location

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    scopes = ['https://www.googleapis.com/auth/drive.readonly']

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'credentials.json', scopes)

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)

        file_id = real_file_id

        # pylint: disable=maybe-no-member
        request = service.files().export_media(fileId=file_id,
                                               mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(F'Download {int(status.progress() * 100)}.')

    except HttpError as error:
        print(F'An error occurred: {error}')
        file = None

    return file.getvalue()


def get_first_free_row(list):
    for i in range(len(list)):
        if str(list[i]).strip() == "nan":
            return i


def detect_changes(prev_last_row, last_row, workbook, conf):
    if (prev_last_row == last_row):
        print("No changes")
    else:
        for i in range(prev_last_row + 1, last_row + 1):
            em = MIMEMultipart("alternative")
            em['From'] = sender
            region = str(workbook.loc[i]['Unnamed: 4']).strip()
            print(region)
            reciever = conf[region]
            em['To'] = reciever
            em['Subject'] = subject
            part = MIMEText(getBody(workbook, i), "html")
            em.attach(part)
            # print(part.as_string)
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
                smtp.login(sender, password)
                smtp.sendmail(sender, reciever, em.as_string())


def getBody(workbook, row):
    info = """
    <html>
    <head>
        <style>
        table, td, th {
            border: 1px solid black; 
            border-collapse: collapse;
            padding: 4px;
            padding-left: 16px;
            padding-right:16px;
        }
        </style>
    </head>
        <body>
            <p>
                Hi there, <br><br>We've received a by-election alert in your patch in case you were not aware. <br> <br>Some info is below: <br><br> <b> </p> <table> <thead>
            <tr>
                <th>
                    Date of election
                </th>
                <th>
                    Authority
                </th>
                <th>
                    Ward
                </th>
                <th>
                    Region
                </th>
                <th>
                    Constituency
                </th>
                <th>
                    Seat Type
                </th>
                <th>
                    Cllr leaving
                </th>
                <th>
                    Party
                </th>
                <th>
                    Reason
                </th>
            </tr>
        </thead>
        <tbody>
                
                <tr>"""
    for i in range(1, 10):
        row_data = str(workbook.loc[row]["Unnamed: " + str(i)])
        if (row_data.strip() == "nan"):
            row_data = " "
        info = info + "<td>" + row_data + " </td> "

    info = info + """ </tr> </tbody></table> <p> </b> <br><br>Best regards, Aiden </p> </body> 
    
    </html>"""
    # print(info)
    return info


if __name__ == '__main__':

    son = open('config.json')
    data = json.load(son)
    son.close()

    # print(data["type"])
    byte_data = export_excel(
        real_file_id='15Tdv4-sXW4yS76zEiiLXI5RrcmPBwKCNihmr564vGQA')
    with open('demo.xlsx', 'wb') as f:
        f.write(byte_data)
        f.close()
    workbook = pd.read_excel('demo.xlsx')
    # for index, row in workbook.iterrows():
    #     print(row["Unnamed: 1"])
    last_row = get_first_free_row(workbook['Unnamed: 1']) - 1
    # print(workbook.loc[last_row])

    if ("last_row" in data.keys()):
        prev_last_row = data["last_row"]
    else:
        prev_last_row = 0

    detect_changes(prev_last_row, last_row, workbook, data)
    data["last_row"] = last_row

    # print(type(data))
    with open('config.json', 'w') as f:
        json.dump(data, f)
        # f.write(str(data))
        f.close()

    #print(workbook.loc[200]['Unnamed: 3'])
