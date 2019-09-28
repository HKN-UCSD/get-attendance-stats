from __future__ import print_function
from datetime import datetime
import utils_parser as up

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
EVENT_LOG_SPREADSHEET_ID = '1vfFN3snw98BcQ0TnWoWHl-Xt2Pdx7IZqbp_KciGVNOg'
FALL_SHEET_NAME = 'Fall 2019'
WINTER_SHEET_NAME = 'Winter 2020'
SPRING_SHEET_NAME = 'Spring 2020' 
RSVP_SIGNIN_COLUMNS = '!B3:L' # Get columns C through L starting on the 3rd row


RANGE_NAMES = [
    FALL_SHEET_NAME + RSVP_SIGNIN_COLUMNS,
    WINTER_SHEET_NAME + RSVP_SIGNIN_COLUMNS
    #SPRING_SHEET_NAME + RSVP_SIGNIN_COLUMNS
]

class Event: 
    def __init__(self):
        self.date = None
        self.name = None
        self.quarter = None
        self.rsvpSpreadsheet = None
        self.signinSpreadsheet = None


def getEvents():

    service = up.getService()

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().batchGet(spreadsheetId=EVENT_LOG_SPREADSHEET_ID, ranges=RANGE_NAMES).execute()
    ranges = result.get('valueRanges', [])
    #print('{0} ranges retrieved.'.format(len(ranges)))

    events = dict()
    if not ranges:
        print('No data found.')
    else:
        for aRange in ranges : 
            if 'values' in aRange.keys() : 
                for row in aRange['values']:
                    #print(row)
                    if len(row) > 10:
                        # Save columns B, C, K, and L, which correspond to indices 0, 1, 9, and 10.
                        newEvent = Event() 
                        newEvent.date = datetime.strptime(row[0], '%m/%d/%Y')
                        newEvent.name = row[1]
                        newEvent.quarter = aRange['range'].split("!")[0]
                        newEvent.rsvpSpreadsheet = row[9]
                        newEvent.signinSpreadsheet = row[10]
                        events[newEvent.name] = newEvent
    return events