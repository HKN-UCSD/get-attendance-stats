from __future__ import print_function
from datetime import datetime
import event_log_parser as elp
import utils_parser as up

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

RANGE = 'Form Responses 1!A2:F'

class Attendee : 
    def __init__(self):
        self.date = None
        self.name = None
        self.email = None
        self.affiliation = None
        self.department = None

def getPastEvents (events, thresholdDate) :
    newEventsDict = dict()
    for eventKey in events.keys(): 
        event = events[eventKey]
        if event.date < thresholdDate: 
            newEventsDict[eventKey] = event
    return newEventsDict

def getFutureEvents (events, thresholdDate) :
    newEventsDict = dict()
    for eventKey in events.keys(): 
        event = events[eventKey]
        if event.date > thresholdDate: 
            newEventsDict[eventKey] = event
    return newEventsDict

numColumnsInRSVPForm = 6
numColumnsInSignupForm = 5

timestampIndex = 0

emailIndex_RSVP = 1
nameIndex_RSVP = 2
affiliation_RSVP = numColumnsInRSVPForm - 1
department_RSVP = affiliation_RSVP - 1 

emailIndex_SignIn = 2
nameIndex_SignIn = 1
affiliation_SignIn = numColumnsInSignupForm - 1
department_SignIn = affiliation_SignIn - 1 

def getAttendees (events) : 
    service = up.getService()
    rsvpAttendees = dict()
    signupAttendees = dict()
    for eventKey in events.keys(): 
        event = events[eventKey]
        eventName = event.name
        rsvpSpreadsheetID = up.getSpreadsheetID(event.rsvpSpreadsheet)

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=rsvpSpreadsheetID,
                                    range=RANGE).execute()
        values = result.get('values', [])

        for value in values: 
            
            # Parse RSVP Form
            if len(value) > numColumnsInRSVPForm - 1:
                attendee = Attendee()
                attendee.date = datetime.strptime(value[timestampIndex], '%m/%d/%Y %H:%M:%S')
                attendee.name = value[nameIndex_RSVP]
                attendee.email = value[emailIndex_RSVP]
                attendee.affiliation = value[affiliation_RSVP]
                attendee.department = value[department_RSVP]
                rsvpAttendees[eventName + "_" + attendee.name] = attendee
            
            # Parse SignIn Form
            if len(value) > numColumnsInSignupForm - 1:
                attendee = Attendee()
                attendee.date = datetime.strptime(value[timestampIndex], '%m/%d/%Y %H:%M:%S')
                attendee.name = value[nameIndex_SignIn]
                attendee.email = value[emailIndex_SignIn]
                attendee.affiliation = value[affiliation_SignIn]
                attendee.department = value[department_SignIn]
                signupAttendees[eventName + "_" + attendee.name] = attendee

    return rsvpAttendees, signupAttendees

def tallyOfficers(attendees) : 
    officersCount = dict()
    officersNames = dict()
    for attendeeKey in attendees.keys() :
        attendee = attendees[attendeeKey]
        if 'Officer' in attendee.affiliation : 
            if attendee.email not in officersCount.keys() :
                officersCount[attendee.email] = 1
                officersNames[attendee.email] = attendee.name
            else :
                officersCount[attendee.email] += 1
    return officersCount, officersNames

# This tallies up the future RSVP count of officers
events = elp.getEvents()
rsvpAttendees, signupAttendees = getAttendees(getFutureEvents(events, datetime.now()))
officerCount, officersNames = tallyOfficers(rsvpAttendees)
for officer in officerCount.keys() : 
    print(officersNames[officer] + " " + str(officerCount[officer]))



