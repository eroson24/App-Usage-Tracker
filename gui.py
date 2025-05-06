import tkinter as tk
from tkinter import ttk
import json
import win32gui as wing
import win32process as winp
import win32com.client as winc
import ctypes
import psutil

def updateDictionary():
    fid.seek(0)
    json.dump(appToCategory, fid, indent=4)
    fid.truncate()

# Create window
window = tk.Tk()
window.title("App Usage Tracker")
window.geometry("700x400")

# Get list of processes
processList = []
processListCOM = winc.GetObject('winmgmts:')
process_ids = [p.ProcessId for p in processListCOM.InstancesOf('win32_process')]
for id in process_ids:
    id = ctypes.c_uint32(id).value
    # Returns either programName.exe or None if on desktop / no active program
    try:
        proc = psutil.Process(id)
        processName = proc.name()
        processList.append(processName)
    except psutil.NoSuchProcess:
        processName = None
# Remove duplicates and empty strings
# Copy list to prevent "Set changed size during iteration" error
processList = set(processList)
processListCopy = processList.copy()
for process in processListCopy:
    if process == "":
        processList.remove(process)

# Import existing app_to_category file
fid = open("app_to_category.json", "r+")
appToCategory = json.load(fid)

# If .exes exist not yet in the file, add them and map to 5. 
for exeName in processList:
    if exeName not in appToCategory:
        appToCategory[exeName] = 5
updateDictionary()

# Begin populating window
# Top text widget
instructions = ttk.Label(master = window, text = "Categorize each .exe below to log its usage into its respective category.", font = "Calibri 12 bold")
instructions.pack(side = 'top', pady = 10)


# Left applications widget
programListFrame = ttk.Frame(master = window)
treeview = ttk.Treeview(master = programListFrame)
for app in processList:
    level1 = treeview.insert("", tk.END, text = app)

programListFrame.pack(side = 'left', padx = 100)
treeview.pack()

# Right options widget
categoryFrame = ttk.Frame(master = window)
selected = tk.IntVar()
r1 = ttk.Radiobutton(categoryFrame, text='Gaming', value=0, variable = selected)
r2 = ttk.Radiobutton(categoryFrame, text='Browsing', value=1, variable = selected)
r3 = ttk.Radiobutton(categoryFrame, text='Programming', value=2, variable = selected)
r4 = ttk.Radiobutton(categoryFrame, text='Music', value=3, variable = selected)
r5 = ttk.Radiobutton(categoryFrame, text='Chatting', value=4, variable = selected)
r6 = ttk.Radiobutton(categoryFrame, text='Other', value=5, variable = selected)

categoryFrame.pack(side = 'right', padx = 100)
r1.pack()
r2.pack()
r3.pack()
r4.pack()
r5.pack()
r6.pack()

# Run window
window.mainloop()