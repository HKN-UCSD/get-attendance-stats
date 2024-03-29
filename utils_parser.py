from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


def getSpreadsheetID(url) :
    # Assumes URL is of the form: 
    # https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/blahblahblah

    indexOfSpreadsheetId = 5 

    parts = url.split("/")
    # Debug code to help visualize the index of the spreadsheet id
    # for i in range(len(parts)) :
    #     print(str(i) + ". " + parts[i])

    if(len(parts) > indexOfSpreadsheetId) : 
        return parts[indexOfSpreadsheetId]
    else : 
        return None


def getService(): 
    # Gets the service needed to read from the spreadsheet. 
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

    service = build('sheets', 'v4', credentials=creds)
    return service