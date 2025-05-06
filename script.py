import os.path
import schedule
import psutil
import time
import win32gui as win
import win32process as winp

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/calendar.events.owned"]

eventStrings = ["playing games", "browsing the web", "programming", "listening to music", "chatting", "on uncategorized apps"]
eventTimeLog = [0, 0, 0, 0, 0, 0]
logCounter = 1
appToCategory = {
    "Battle.net.exe": "Gaming",
    "Overwatch.exe": "Gaming",
    "steam.exe": "Gaming",
    "msedge.exe": "Browsing",
    "chrome.exe": "Browsing",
    "Code.exe": "Programming",
    "Spotify.exe": "Music",
    "Discord.exe": "Chatting"
}


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

def createEvent():
    global timeBegin
    global timeEnd
    global logCounter

    # Generate an event description string
    eventDescription = ""
    i = 0
    for a in eventTimeLog:
        if a != 0:
            eventDescription = eventDescription + (f"Time spent {eventStrings[i]}: {a} seconds\n")
        i += 1

    timeEnd = time.strftime("%Y-%m-%dT%H:%M:00%z")[:-2] + ":" + time.strftime("%Y-%m-%dT%H:%M:00%z")[-2:]
    try:
        service = build("calendar", 'v3', credentials=creds)
        event = {
                'summary': f'Log #{logCounter}',
                'description': eventDescription,
                'colorId': '1',
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
        event = service.events().insert(calendarId='primary', body=event).execute()
        
    except HttpError as error: 
        print("An error occurred :( - ", error)

    # reset timeBegin and eventTimeLog, andincrement logCounter
    timeBegin = time.strftime("%Y-%m-%dT%H:%M:00%z")[:-2] + ":" + time.strftime("%Y-%m-%dT%H:%M:00%z")[-2:]
    for i in range(len(eventTimeLog)):
        eventTimeLog[i] = 0
    logCounter += 1

def logActivity():
    global eventTimeLog
    global appToCategory

    #  Get the name of the program the user is on
    activeWindow = win.GetForegroundWindow()
    pid = winp.GetWindowThreadProcessId(activeWindow)
    pid = pid[1]

    # Returns either programName.exe or None if on desktop / no active program
    try:
        proc = psutil.Process(pid)
        processName = proc.name()
    except psutil.NoSuchProcess:
        processName = None

    # Try to locate type of event and log
    uncategorized = True

    for key in appToCategory:
        if key == processName:
            uncategorized = False
            typeOfProgram = appToCategory[key]
            if typeOfProgram == "Gaming":
                eventTimeLog[0] += 1
            if typeOfProgram == "Browsing":
                eventTimeLog[1] += 1
            if typeOfProgram == "Programming":
                eventTimeLog[2] += 1
            if typeOfProgram == "Music":
                eventTimeLog[3] += 1
            if typeOfProgram == "Chatting":
                eventTimeLog[4] += 1

    if uncategorized:
        eventTimeLog[5] += 1

auth()

# Set up scheduler
schedule.every(4).minutes.do(createEvent)
schedule.every(1).second.do(logActivity)

# Initial timeBegin and timeEnd
timeBegin = time.strftime("%Y-%m-%dT%H:%M:00%z")[:-2] + ":" + time.strftime("%Y-%m-%dT%H:%M:00%z")[-2:]
timeEnd = None

while True:
    # Check for events every second
    schedule.run_pending()
    time.sleep(1)
    
    