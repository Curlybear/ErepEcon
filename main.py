import gspread
import configparser

import requests
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import json
import os
import time

def getDay():
    start_date = datetime.datetime(day=21, month=11, year=2007)
    erepDay = datetime.datetime.now() - start_date
    return erepDay.days


def initSheet(stateSheet, orgsSheet):
    # Setup current day
    stateSheet.update_acell('Z1', "Last update erep day")
    stateSheet.update_acell('Z2', str(getDay()))

    # Setup current row
    stateSheet.update_acell('Y1', "Current Row")
    stateSheet.update_acell('Y2', 2)

    # Setup stateSheet headers
    stateSheet.update_acell('A1', 'Day')
    stateSheet.update_acell('B1', 'State CC')
    stateSheet.update_acell('C1', 'State Gold')
    stateSheet.update_acell('D1', 'Total CC with orgs')
    stateSheet.update_acell('E1', 'Total Gold with orgs')
    stateSheet.update_acell('F1', 'Description')

    # Setup orgsSheet headers
    orgsSheet.update_acell('A1', 'Day')
    i = 2
    orgs = json.loads(config.get("DEFAULT", "orgs"))
    orgStacks = [orgs[n:n + 10] for n in range(0, len(orgs), 10)]

    for orgStack in orgStacks:
        r = requests.get(
            'https://api.erepublik-deutschland.de/' + apiKey + '/organizations/details/' + ','.join(orgStack))

        if r.headers['X-Rate-Limit-Remaining'] == '0':
            time.sleep(float(r.headers['X-Rate-Limit-Reset']) + 5)

        obj = json.loads(r.text)
        for org in orgStack:
            orgDetail = obj['organizations'][str(org)]
            orgsSheet.update_cell(1, i, orgDetail['name'] + ' CC')
            i += 1
            orgsSheet.update_cell(1, i, orgDetail['name'] + ' Gold')
            i += 1
    orgsSheet.update_cell(1, i, 'Total CC')
    i += 1
    orgsSheet.update_cell(1, i, 'Total Gold')

    return 2


def checkRun(stateSheet):
    if stateSheet.acell('A1').value:
        sheetDay = stateSheet.acell('Z2').value
        currentDay = getDay()
        if int(sheetDay) == int(currentDay):
            exit(1)
    else:
        return 1


def updateSheet(stateSheet):
    # Get Value
    row = stateSheet.acell('Y2').value
    # Update Value
    row = int(row) + 1
    stateSheet.update_acell('Y2', row)
    stateSheet.update_acell('Z2', str(getDay()))
    return row


def fetchData(stateSheet, orgsSheet, currentRow):
    stateSheet.update_cell(currentRow, 1, getDay())
    orgsSheet.update_cell(currentRow, 1, getDay())

    countryId = str(config['DEFAULT']['country_id'])

    r = requests.get('https://api.erepublik-deutschland.de/' + apiKey + '/countries/details/' + countryId)
    obj = json.loads(r.text)
    stateSheet.update_cell(currentRow, 2, obj['countries'][countryId]['economy']['cc'])
    stateSheet.update_cell(currentRow, 3, obj['countries'][countryId]['economy']['gold'])

    i = 2
    orgs = json.loads(config.get("DEFAULT", "orgs"))
    orgStacks = [orgs[n:n + 10] for n in range(0, len(orgs), 10)]

    for orgStack in orgStacks:
        r = requests.get(
            'https://api.erepublik-deutschland.de/' + apiKey + '/organizations/details/' + ','.join(orgStack))

        if r.headers['X-Rate-Limit-Remaining'] == '0':
            time.sleep(float(r.headers['X-Rate-Limit-Reset']) + 5)

        obj = json.loads(r.text)
        for org in orgStack:
            orgDetail = obj['organizations'][str(org)]
            orgsSheet.update_cell(currentRow, i, orgDetail['money']['account']['cc'])
            i += 1
            orgsSheet.update_cell(currentRow, i, orgDetail['money']['account']['gold'])
            i += 1

    orgsSheet.update_cell(currentRow, len(orgs)*2 + 2, '=' + (''.join(chr(65 + col) + str(currentRow) + '+' for col in range(1, len(orgs)*2, 2)))[:-1])
    orgsSheet.update_cell(currentRow, len(orgs)*2 + 3, '=' + (''.join(chr(65 + col) + str(currentRow) + '+' for col in range(2, len(orgs)*2 + 1, 2)))[:-1])

    stateSheet.update_cell(currentRow, 4, '=Sheet2!' + chr(65 + len(orgs)*2 + 1) + str(currentRow) + '+B'+ str(currentRow))
    stateSheet.update_cell(currentRow, 5, '=Sheet2!' + chr(65 + len(orgs)*2 + 2) + str(currentRow) + '+C'+ str(currentRow))

# Config reader
dir = os.path.dirname(__file__)
filename = os.path.join(dir, 'config.ini')
config = configparser.ConfigParser()
config.read(filename)

# API Key
apiKey = config['DEFAULT']['api_key']

# Setup gSpread
scope = ['https://spreadsheets.google.com/feeds']
jsonFilename = config['DEFAULT']['google_key']
dir = os.path.dirname(__file__)
filename = os.path.join(dir, jsonFilename)
credentials = ServiceAccountCredentials.from_json_keyfile_name(filename, scope)
gc = gspread.authorize(credentials)
sheetKey = config['DEFAULT']['sheet_key']
stateSheet = gc.open_by_key(sheetKey).get_worksheet(0)
orgsSheet = gc.open_by_key(sheetKey).get_worksheet(1)

if checkRun(stateSheet):
    currentRow = initSheet(stateSheet, orgsSheet)
else:
    currentRow = updateSheet(stateSheet)

fetchData(stateSheet, orgsSheet, currentRow)
