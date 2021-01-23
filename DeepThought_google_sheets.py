# -*- coding: utf-8 -*-
"""
Created on Sat Jan 23 19:25:17 2021

@author: Lenovo
"""
#client id: 196568370672-0nv9ttle82b0gmlpse1hg9rcove15boh.apps.googleusercontent.com
#client secret: i5N0L5pJ6c_5i5GHNj8w7l9p
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
#SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_SPREADSHEET_ID = '1K1Jx5DPIhyu066ClVZsrcyLXONH6LoOChIVPtbyt-zY'

SAMPLE_RANGE_NAME = 'Sheet1!A1:I12'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    global service
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('Name, Major:')
        for row in values:
            print('| %s | %s | %s | %s | %s | %s | %s | %s | %s |' % (row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8], ))#,row[3],row[4],row[5],row[6],row[7],row[8]))


#write values into the sheets
        response_date = service.spreadsheets().values().update(
        spreadsheetId='1RvgoMeeJOXLAxZ0MIoK6lTnSP0YN9VPpeYgRu9gyE0o',
        valueInputOption='RAW',
        range=SAMPLE_RANGE_NAME,
        body=dict(
            majorDimension='ROWS',
            values=values)
        ).execute()
        print('Sheet successfully Updated')
    
if __name__ == '__main__':
    main()
