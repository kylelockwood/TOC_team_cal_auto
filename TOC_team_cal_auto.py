#! python3
# Uses data from team schedule worksheet and creates calendar events on the TOC site

import pyautogui, openpyxl
from datetime import timedelta
from ics import Calendar, Event

"""
class Unbuffered(object):
   def __init__(self, stream):
       self.stream = stream
   def write(self, data):
       self.stream.write(data)
       self.stream.flush()
   def writelines(self, datas):
       self.stream.writelines(datas)
       self.stream.flush()
   def __getattr__(self, attr):
       return getattr(self.stream, attr)

import sys
sys.stdout = Unbuffered(sys.stdout)
"""

def get_xl_data(sheet):
    print(f'Parsing calendar data... ', flush=True, end='')
    # Return the dates and schedule data
    dates = []
    # Get dates
    for r in range(1, 40, 3):
        dates.append(sheet.cell(row=r, column=1).value)
    descriptions = []
    # Get schedule data
    for r in range(2, 41, 3):
        description = ''
        for c in range(1,40):
            description += str(sheet.cell(row=r, column=c).value) + ' '
        descriptions.append(description)
    # Create dictionaries
    calDataList = []
    calData = {}
    for i in range(len(dates)):
        name = 'Team Schedule'
        # Add 'Empty Services'
        if dates[i] == None:
            name = 'Empty Service'
            dates[i] = dates[i-1] + timedelta(days=7) # Add the date back in, 7 days after the previous
            descriptions[i] = name
        calData['name'] = name
        calData['date'] = dates[i]
        calData['description'] = descriptions[i]
        # Create list of dictionaries
        calDataList.append(calData.copy())
    print('Done')
    return calDataList


def create_ics(calData, outFile):
    cal = Calendar()
    event = Event()
    print(f'Writing ICS file \'{outFile}.ics\'... ', flush=True, end='')
    with open(outFile + '.ics', 'w') as f:
        for data in calData:
            event.name = data.get('name')
            event.begin = data.get('date')
            event.description = data.get('description')
            event.location = 'The Oregon Community 700 NE Dekum St. Portland OR'
            cal.events.add(event)
            f.writelines(cal)
    print('Done')
    return


# Load excel sheet
path = 'E:\\Google Drive\\TOC\\'
fileName = 'TOC Team Schedule Jan-Mar20.xlsm'
print(f'Reading {fileName}... ', flush=True, end='')
workbook = path + fileName
wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True) 
ws = wb['Web Cal']
print('Done')


# Get schedule data from excel file
calData = get_xl_data(ws)

# Create ICS file from schedule data
create_ics(calData, 'TOC_clover_cal')


# TODO open calendar site in browser, full screen
# TODO check to see the page has loaded
# TODO click through to the calendar
# TODO Import ICS file
# TODO Profit?