from __future__ import print_function
from collections import OrderedDict
from dateutil import relativedelta
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common import exceptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from time import sleep
import csv
import datetime
import os
import os.path
import pickle
import selenium
import time

# # If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

## input vars
name =  # input for calendar events.
username =  # fill with string
password =  # fill with string
url =  # URL to login portal, string
calendarId =  # fill with string
## notification in minutes
downloadDir = r'D:\dev\ortec_downloads'
downloadedFile = r'D:/dev/ortec_downloads/Employee Schedule ESS.csv' # Filname could be static
chromeDriverPath = r'D:\dev\chromedriver.exe'
notificationTime = 60
sleep_time = 2
num_retries = 4
timeZone = 'Europe/Amsterdam'

## set chrome options
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : downloadDir}
chrome_options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=chromeDriverPath)


##gcal stuff
def createEvent(payload):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    event = service.events().insert(calendarId=calendarId, body=payload).execute()
    print(payload['summary'] +' ' + knownStartDate + ' created: %s' % (event.get('htmlLink')))

# checks if an event exists in gcal
def checkEventExists(payload):
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    events_result = service.events().list(calendarId=calendarId, timeMin=payload['start'].get('dateTime'),
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])
    knownEvent = False
    if not events:
        print('No upcoming events found.')
    for event in events:
        knownStartDate =  event['start'].get('dateTime')
        payloadStartDate = payload['start'].get('dateTime')
        knownEndDate =  event['end'].get('dateTime')
        payloadEndDate = payload['end'].get('dateTime')
        knownEventName = payload['summary']
        payloadEventName = event['summary']
        if knownStartDate == payloadStartDate and knownEventName == payloadEventName and knownEndDate == payloadEndDate:
            print(event['summary'] +' ' + knownStartDate + ' aldready exists, skipping')
            knownEvent = True
    if knownEvent == True:
        return True
    else:
        return False


def pushtoGoogle():
    if os.path.exists(downloadedFile):
        workEvents = []
        with open(downloadedFile) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            for row in csv_reader:
                if row and line_count >= 2 and row[4]:
                    date = row [3]
                    location = row[4]
                    time = row[5]
                    code = row[6]
                    line_count += 1
                    parsedEvent = {
                        "Date": date,
                        "Location": location,
                        "Time": time,
                        "Code": code
                    }
                    workEvents.append(parsedEvent)
                else:
                    line_count += 1

        dedupWorkEvents = []
        for event in workEvents:
            if event not in dedupWorkEvents:
                dedupWorkEvents.append(event)

        for event in dedupWorkEvents:
            time = event.get("Time").split('-')
            startTime = time[0].strip()
            endTime = time[1].strip()
            startDateTime = datetime.datetime.strptime(event.get("Date") + " " + startTime, '%m/%d/%Y %H:%M').strftime("%Y-%m-%dT%H:%M:%S+02:00")
            endDateTime = datetime.datetime.strptime(event.get("Date") + " " + endTime, '%m/%d/%Y %H:%M').strftime("%Y-%m-%dT%H:%M:%S+02:00")
            payload = {
                'summary': name +" Dienst " + event.get("Code"),
                'location': event.get("Location"),
                'description': event.get("Code"),
                'start': {
                    'dateTime': startDateTime,
                    'timeZone': timeZone,
                },
                'end': {
                    'dateTime': endDateTime,
                    'timeZone': timeZone,
                },
                'recurrence': [],
                'attendees': [],
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    {'method': 'popup', 'minutes': notificationTime},
                    ],
                },
                }
            if checkEventExists(payload) == False:
                createEvent(payload)

## set date vars
todayStr = datetime.datetime.now().strftime("%m") + '/' + \
    datetime.datetime.now().strftime("%d") + '/' + \
    datetime.datetime.now().strftime("%Y") + ' 12:00:00 AM'
nextMonth = datetime.date.today() + relativedelta.relativedelta(months=1)
nextMonthStr = nextMonth.strftime("%m") + '/' + \
    nextMonth.strftime("%d") + '/' + \
    nextMonth.strftime("%Y") + ' 12:00:00 AM'
monthAfterNext = nextMonth  + relativedelta.relativedelta(months=1)
monthAfterNextStr = monthAfterNext.strftime("%m") + '/' + \
    monthAfterNext.strftime("%d") + '/' + \
    monthAfterNext.strftime("%Y") + ' 12:00:00 AM'


driver.maximize_window()
driver.get(url)
usernameField = driver.find_element_by_id('ctl00_ContentPlaceHolder1_Username')
usernameField.send_keys(username)
passwordField = driver.find_element_by_id('ctl00_ContentPlaceHolder1_Password')
passwordField.send_keys(password)
driver.find_element_by_name('ctl00$ContentPlaceHolder1$Button1').click()
sleep(sleep_time)


openIframe = len(driver.find_elements_by_tag_name('iframe'))
while openIframe == 0:
    try:
        openIframe = len(driver.find_elements_by_tag_name('iframe'))
        driver.find_element_by_xpath("//*[contains(text(), 'Reports')]").click()
        driver.find_element_by_xpath("//*[contains(text(), 'Personal schedule')]").click()
        time.sleep(sleep_time)
    except:
        print("Can't open reports iframe. Retrying...")
        time.sleep(0.2)

# ########Section-1

# get the list of iframes present on the web page using tag "iframe"

seq = driver.find_elements_by_tag_name('iframe')

print("No of frames present in the web page are: ", len(seq))
sleep(sleep_time)

# #switching between the iframes based on index

for index in range(len(seq)):

    driver.switch_to.default_content()

    iframe = driver.find_elements_by_tag_name('iframe')[index]
    driver.switch_to.frame(iframe)

    driver.implicitly_wait(5)

    untilField = driver.find_element_by_id('ReportViewer_ctl04_ctl05_txtValue')
    untilField.clear()
    sleep(sleep_time)
    untilField.send_keys(nextMonthStr)
    # open report
    driver.find_element_by_id('ReportViewer_ctl04_ctl00').click()
    time.sleep(sleep_time)
    if os.path.exists(downloadedFile):
        os.remove(downloadedFile)

    while not os.path.exists(downloadedFile):
        print("Trying to get current month report")
        driver.find_element_by_id('ReportViewer_ctl06_ctl04_ctl00_ButtonImg').click()
        driver.find_element_by_xpath("//*[contains(text(), 'CSV')]").click()
        time.sleep(sleep_time)

    time.sleep(sleep_time)
    pushtoGoogle()

    if os.path.exists(downloadedFile):
        os.remove(downloadedFile)

    fromField = driver.find_element_by_id('ReportViewer_ctl04_ctl03_txtValue')
    fromField.clear()
    sleep(sleep_time)
    fromField.send_keys(nextMonthStr)
    sleep(sleep_time)
    driver.find_element_by_id('ReportViewer_ctl04_ctl00').click()
    sleep(sleep_time)
    driver.implicitly_wait(5)
    untilField = driver.find_element_by_id('ReportViewer_ctl04_ctl05_txtValue')
    untilField.clear()
    sleep(sleep_time)
    untilField.send_keys(monthAfterNextStr)
    sleep(sleep_time)
    driver.find_element_by_id('ReportViewer_ctl04_ctl00').click()
    sleep(sleep_time)
    while not os.path.exists(downloadedFile):
        print("Trying to get next month's report")
        driver.find_element_by_id('ReportViewer_ctl06_ctl04_ctl00_ButtonImg').click()
        driver.find_element_by_xpath("//*[contains(text(), 'CSV')]").click()
        sleep(sleep_time)

    sleep(sleep_time)
    pushtoGoogle()