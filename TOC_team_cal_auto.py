#! python3
# Uses data from team schedule worksheet and creates calendar events on the TOC site

import pyautogui, openpyxl
from datetime import timedelta


def get_xl_data(sheet):
    # Return the dates and schedule data
    dates = []
    # Get dates
    for r in range(1, 40, 3):
        # TODO Reformat dates?
        dates.append(str(sheet.cell(row=r, column=1).value).split(' ')[0]) # May not want this as a string depending on calendar API data type
    datas = []
    # Get schedule data
    for r in range(2, 41, 3):
        data = ''
        for c in range(1,40):
            data += str(sheet.cell(row=r, column=c).value) + ' '
        datas.append(data)
    # Create dictionaries
    calDataList = []
    calData = {}
    for i in range(len(dates)):
        # Add 'Empty Services'
        if dates[i] == None:
            dates[i] = dates[i-1] + timedelta(days=7)
            datas[i] = 'Empty Service'
        calData['date'] = dates[i]
        calData['data'] = datas[i]
        calData['i'] = i
        # Create list of dictionaries
        calDataList.append(calData.copy())
    return calDataList



# Get relevant excel data from sheet
path = 'E:\\Google Drive\\TOC\\'
workbook = path + 'TOC Team Schedule Jan-Mar20.xlsm'
wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True) 
ws = wb['Web Cal']

calData = get_xl_data(ws)
print(calData)

# TODO Does the clover calendar have a publically accessable API?

# TODO if no API then:
    # TODO open calendar site in browser, full screen
    # TODO check to see the page has loaded
    # TODO click through to the calendar
    # TODO Find dates
    # TODO create new events
    # TODO fill out the event data