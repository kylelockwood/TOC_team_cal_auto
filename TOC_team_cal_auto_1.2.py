#! python3
"""
Uses data from team schedule worksheet and creates calendar events on the google calendar and TOC site
"""

from __future__ import print_function
from datetime import timedelta
import pickle
import os.path
import sys
import openpyxl
from ics import Calendar, Event
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']
SCRIPTPATH = os.path.dirname(os.path.realpath(sys.argv[0])) + '\\'


def main():
    # Load excel sheet
    path = '\\\\192.168.86.214\\SharePi\\' # Network drives need double slashes \\ at the top
    fileName = 'TOC Team Schedule Jan-Mar20.xlsm'
    worksheet = load_workbook(path, fileName)['Web Cal']

    # Get schedule data from excel file
    calData = get_xl_data(worksheet)

    # Connect to the Google Calendar API and get calendar names and IDs
    service = check_creds(SCRIPTPATH)
    calids = get_cal_ids(service)
    calName = ['TOC test']

    # Upload calendar data to Google Calendar
    upload_to_gcal(calData, service, calName, calids)

    # Create .ics file (unnecessary at the moment, but may have a future use)
    # create_ics(calData, 'TOC_cal')


def load_workbook(path, fileName):
    """
    Loads the excel workbook for reading
    """
    print(f'Reading \'{fileName}\'... ', flush=True, end='')
    workbook = openpyxl.load_workbook(path + fileName, read_only=True, data_only=True) 
    #print('Done')
    return workbook


def get_xl_data(sheet):
    """
    Returns list of dictionaries: name, date and description from the worsheet
    """
    #print(f'Parsing calendar data... ', flush=True, end='')
    dates = []
    # Get dates
    for r in range(1, 40, 3):
        dates.append(sheet.cell(row=r, column=1).value)
    descriptions = []
    # Get description data
    for r in range(2, 41, 3):
        description = ''
        for c in range(1, 40):
            description += str(sheet.cell(row=r, column=c).value) + ' '
        descriptions.append(description)
    # Create dictionaries
    calDataList = []
    calData = {}
    for i in range(len(dates)):
        name = 'Team Schedule'
        # Add 'Empty Services'
        if dates[i] is None:
            name = 'Empty Service'
            dates[i] = dates[i-1] + timedelta(days=7) # Add the date back in, 7 days after the previous
            descriptions[i] = name
        calData['name'] = name
        calData['date'] = dates[i].date()
        calData['description'] = descriptions[i]
        # Create list of dictionaries
        calDataList.append(calData.copy())
    print('Done')
    return calDataList


def create_ics(calData, outFile):
    """
    Creates an .ics file for uploading to calendar services
    """
    cal = Calendar()
    print(f'Writing ICS file \'{outFile}.ics\'... ', flush=True, end='')
    for data in calData:
        event = Event()
        event.name = data.get('name')
        event.begin = data.get('date')
        event.make_all_day()
        event.description = data.get('description')
        event.location = 'The Oregon Community 700 NE Dekum St. Portland OR'
        cal.events.add(event)
    with open(outFile + '.ics', 'w', newline='') as f: # Clover calendar wont read ics with extra carriage returns
        f.writelines(cal)
    print('Done')
    return


def check_creds(credPath):
    """
    Opens Google Calendar API connection
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
                credPath + 'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    return service


def get_cal_ids(service):
    """
    Return all calendar names (key) and ids (value) associated with the master calendar email address
    """
    page_token = None
    calids = {}
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            calids[calendar_list_entry['summary']] = calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            return calids


def upload_to_gcal(calData, service, calNames, calids):
    """
    Uploads event data to the calendars passed
    """
    for calendar in calNames:
        for cal in calids:
            if cal == calendar:
                calid = calids.get(cal)
                print(f'Loading events to calendar \'{cal}\'...')
                for data in calData:
                    name = data.get('name')
                    startDate = data.get('date').strftime('%Y-%m-%d')
                    endDate = (data.get('date') + timedelta(days=1)).strftime('%Y-%m-%d')
                    description = data.get('description')
                    # Create the event
                    event = {
                        'summary': name,
                        'location': 'The Oregon Community 700 NE Dekum St. Portland, OR 97211',
                        'description': description,
                        'start': {
                            'date': startDate,
                            'timeZone': 'America/Los_Angeles',
                        },
                        'end': {
                            'date': endDate,
                            'timeZone': 'America/Los_Angeles',
                        }
                    }
                    event = service.events().insert(calendarId=calid, body=event).execute()
                    print(f'     Event \'{name}\' created on {startDate}')
                print('Done')
        if calendar not in calids:
            print(f'Calendar \'{calendar}\' not found.')
    return


if __name__ == '__main__':
    main()


# Client ID
# 551233068350-dk9d09b6roi1fiiipbmq6racdfo4cb6o.apps.googleusercontent.com

# Client Secret
# lwj98MaanPxUnw3mvixO1dKt