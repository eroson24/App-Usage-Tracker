import ctypes
import json
import os
import schedule
import psutil
import subprocess
import time
import win32gui as win
import win32process as winp

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/calendar.events.owned"]

# Run gui.py to update appToCategory dictionary before running rest of script.py
subprocess.run(["python", "gui.py"])
time.sleep(2)  # short pause before tracking starts

eventStrings = ["playing games", "browsing the web", "programming", "listening to music", "chatting", "on uncategorized apps"]
eventTimeLog = [0, 0, 0, 0, 0, 0]
eventCounter = 1
with open("app_to_category.json", "r") as dictionary:
    appToCategory = json.load(dictionary)

def auth():
    # Google Calendar authentication
    global creds
    global service
    
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json")
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

    # Save the credentials for the next run.
    with open("token.json", "w") as token:
      token.write(creds.to_json())

    # Build calendar service
    try:
        service = build("calendar", 'v3', credentials=creds)
    except HttpError as error: 
        print("An error occurred :( - ", error)

def createEvent():
    global timeBegin
    global eventCounter
    global createEventInterval

    # Generate an event description string
    eventDescription = ""
    
    i = 0
    for a in eventTimeLog:
        if a != 0:
            eventDescription += (f"Time spent {eventStrings[i]}: {a // 60} min {a % 60} seconds ({format(100 * (a / (createEventInterval * 60)), '.2f')}%)\n")
        i += 1
    eventDescription += "\n"


    timeEnd = time.strftime("%Y-%m-%dT%H:%M:00%z")[:-2] + ":" + time.strftime("%Y-%m-%dT%H:%M:00%z")[-2:]
    try:
        event = {
                'summary': f'Log #{eventCounter}',
                'description': eventDescription,
                'colorId': (eventCounter % 11) + 1,
                'start': {
                'dateTime': timeBegin
                },
                'end': {
                'dateTime': timeEnd
                },
                'reminders': {
                'useDefault': False,
                }
                }
        # Add event to primary calendar
        service.events().insert(calendarId='primary', body=event).execute()
        
    except HttpError as error: 
        print("An error occurred :( - ", error)

    # reset timeBegin and eventTimeLog, andincrement eventCounter
    timeBegin = time.strftime("%Y-%m-%dT%H:%M:00%z")[:-2] + ":" + time.strftime("%Y-%m-%dT%H:%M:00%z")[-2:]
    for i in range(len(eventTimeLog)):
        eventTimeLog[i] = 0
    eventCounter += 1

def logActivity():
    global eventTimeLog
    global appToCategory

    #  Get the name of the program the user is on
    activeWindow = win.GetForegroundWindow()
    pid = winp.GetWindowThreadProcessId(activeWindow)
    # Handles error of pid being treated as negative. Converts to uint32
    pid = ctypes.c_uint32(pid[1]).value

    # Returns either programName.exe or None if on desktop / no active program
    try:
        proc = psutil.Process(pid)
        processName = proc.name()
    except psutil.NoSuchProcess:
        processName = None

    # Locate type of event and log
    if processName in appToCategory:
        typeOfProgram = appToCategory[processName]
    else:
        typeOfProgram = 5
    eventTimeLog[typeOfProgram] += 1

auth()

# Set up scheduler
createEventInterval = 5
schedule.every(createEventInterval).minutes.do(createEvent)
schedule.every(1).second.do(logActivity)

# Initialize timeBegin
timeBegin = time.strftime("%Y-%m-%dT%H:%M:00%z")[:-2] + ":" + time.strftime("%Y-%m-%dT%H:%M:00%z")[-2:]

while True:
    # Check for events every second
    schedule.run_pending()
    time.sleep(1)
    
    