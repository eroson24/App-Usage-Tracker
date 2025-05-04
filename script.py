import datetime as dt
import os.path
import sched
import time
import win32gui as win

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/calendar.events.owned"]

def checkProgram():
    global cacheProgram
    global currentProgram
    
    activeWindow = win.GetForegroundWindow()
    currentProgram = win.GetWindowText(activeWindow)
    
    if (cacheProgram == currentProgram):
        print("Same program! You are on " + currentProgram)
    else:
        print("Program switch detected! ")
        print("From " + cacheProgram + " to " + currentProgram)
    
    cacheProgram = currentProgram

def auth():
    # Google Calendar authentication
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
    try:
        service = build("calendar", 'v3', credentials=creds)
        created_event = service.events().quickAdd(
        calendarId='primary',
        text='Appointment on June 3rd 10am-10:25am').execute()
    except HttpError as error: 
        print("An error occurred :( - ", error)

auth()
cacheProgram = "None"
currentProgram = None

activeWindow = win.GetForegroundWindow()
currentProgram = win.GetWindowText(activeWindow)

checkProgramScheduler = sched.scheduler(time.time, time.sleep)

while True:
    checkProgramScheduler.enter(delay = 60, priority = 1, action = checkProgram)
    checkProgramScheduler.run()