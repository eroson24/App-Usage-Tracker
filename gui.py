import tkinter as tk
from tkinter import ttk
import win32gui as wing
import win32process as winp
import win32com.client as winc
import ctypes
import psutil

# Create window
window = tk.Tk()
window.title("App Usage Tracker")
window.geometry("500x300")

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
# Remove duplicates
processList = set(processList)

# Populate window
instructions = ttk.Label(master = window, text = "Categorize each .exe below to log its usage into its respective category.", font = "Calibri 12")
instructions.pack()

# Left widget
programListFrame = ttk.Frame(master = window)
treeview = ttk.Treeview(master = programListFrame)
for app in processList:
    level1 = treeview.insert("", tk.END, text = app)

programListFrame.pack()
treeview.pack()

# Right widget
categoryFrame = ttk.Frame(master = window)
selected = tk.StringVar()
r1 = ttk.Radiobutton(categoryFrame, text='Gaming', value="0", variable = selected)
r2 = ttk.Radiobutton(categoryFrame, text='Browsing', value="1", variable = selected)
r3 = ttk.Radiobutton(categoryFrame, text='Programming', value="2", variable = selected)
r4 = ttk.Radiobutton(categoryFrame, text='Music', value="3", variable = selected)
r5 = ttk.Radiobutton(categoryFrame, text='Chatting', value="4", variable = selected)
r6 = ttk.Radiobutton(categoryFrame, text='Other', value="5", variable = selected)

categoryFrame.pack()
r1.pack()
r2.pack()
r3.pack()
r4.pack()
r5.pack()
r6.pack()

# Run window
window.mainloop()