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

#Alphabetize and update dictionary
appToCategory = dict(sorted(appToCategory.items()))
updateDictionary()

# Begin populating window
# Top text widget
instructions = ttk.Label(master = window, text = "Categorize each .exe below to log its usage into its respective category.", font = "Calibri 12 bold")
instructions.pack(side = 'top', pady = 10)

# Change detecting functions
# Detect tree change
def selectedExeChange():
    selected_item = treeview.selection()
    # Only proceed if something is selected
    if selected_item:  
        exe_name = treeview.item(selected_item[0])["text"]
        selected.set(appToCategory[exe_name])

# Detect radio button change
def selectedButtonChange():
    selected_item = treeview.selection()
    # Only proceed if something is selected
    if selected_item: 
        exe_name = treeview.item(selected_item[0])["text"]
        selectedInt = selected.get()
        appToCategory[exe_name] = selectedInt
    updateDictionary()
    

# Left applications widget
programListFrame = ttk.Frame(master = window)
treeview = ttk.Treeview(master = programListFrame, selectmode="browse")
for app in appToCategory:
    treeview.insert("", tk.END, text = app)
programListFrame.pack(side = 'left', padx = 100, pady = 60, fill='both', expand=True)
treeview.pack(fill='both', expand=True)
treeview.bind("<<TreeviewSelect>>", lambda e: selectedExeChange())

# Right options widget
categoryFrame = ttk.Frame(master = window)
selected = tk.IntVar()
r1 = ttk.Radiobutton(categoryFrame, text='Gaming', value=0, variable = selected, command=selectedButtonChange)
r2 = ttk.Radiobutton(categoryFrame, text='Browsing', value=1, variable = selected, command=selectedButtonChange)
r3 = ttk.Radiobutton(categoryFrame, text='Programming', value=2, variable = selected, command=selectedButtonChange)
r4 = ttk.Radiobutton(categoryFrame, text='Music', value=3, variable = selected, command=selectedButtonChange)
r5 = ttk.Radiobutton(categoryFrame, text='Chatting', value=4, variable = selected, command=selectedButtonChange)
r6 = ttk.Radiobutton(categoryFrame, text='Other', value=5, variable = selected, command=selectedButtonChange)

categoryFrame.pack(side = 'right', padx = (0,100))
r1.pack(anchor='w')
r2.pack(anchor='w')
r3.pack(anchor='w')
r4.pack(anchor='w')
r5.pack(anchor='w')
r6.pack(anchor='w')

# Run window
window.mainloop()