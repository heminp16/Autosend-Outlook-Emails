from pathlib import Path #need
import os, sys, re, ctypes
import PySimpleGUI as sg
import pandas as pd
import tkinter as tk
from tkinter import ttk
import win32com.client as win32

# Main Codes
icon0= "favicon.ico"

def index_containing_substring(list, substring):
        for i, s in enumerate(list):
            if substring in s:
                return i
        return -1

def extractExcel(excel_path):
    df= pd.read_excel(excel_path).dropna(how="all")  #Dropping entire rows with NA

    if "Worker" in df.columns: #Checks to see if Col. "Worker" is in Excel file. If not, throws an error.
        df["Worker"]= df["Worker"].astype("string") #Converting Worker name to Str
        WorkersList= df["Worker Email"].dropna().to_list() #Extracting Worker Emails, and dropping NA (the Applied Filter row) 

        # Extracting Week Range from the "Applied Filter" cell block
        weekRange= df.iloc[-1,0] # Selecting last row
        split= weekRange.split("\n") #Splitting the Applied Filter row to sections

        # Code chunk to figure out where the Week Range is (Index)
        substring= "Week Range is"
        indexLocation = index_containing_substring(split, substring) #prints out the index number of where "Week Range is" is located 

        rangeStr= split[indexLocation] # Selected the Week Range  #'Week Range is x/xx/20xx - x/xx/20xx'

        dateSubject= rangeStr.split()[3:] # Selecting only the Dates # ['x/xx/20xx', '-', 'x/xx/20xx']
        dateSubject= " ".join(dateSubject) # Concatenate previous list to create one str 'x/xx/20xx - x/xx/20xx'
        return dateSubject, WorkersList  
    else:
        errCode = 1
        sg.popup_error("This might be the incorrect Excel file. Columns: 'Worker' or/and 'Worker Email' are missing", icon=Logo)
        return errCode

# Get the computer username of the user. Swap their username to the correct order: First_Name Last_Name
def getSignatureName():
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3
 
    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)
 
    nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(NameDisplay, nameBuffer, size)
    username= nameBuffer.value # Last_Name, First_Name
    return username

username= getSignatureName()
signature= username.replace(",", "").split() #remove the , (comma)
signature= (signature[1], signature[0]) #flips the name ["First_Name", "Last_Name"]
signature= " ".join(signature) # Concatenate list to create one str First_Name Last_Name

# Reads Subject Name
def readSubjectName():
    file_name = "emailBody.txt"
    file_read = open(file_name, "r")
    text = "{ENTER SUBJECT AFTER THIS}"
    lines = file_read.readlines()
    new_list = []
    idx = 0
    for line in lines:
        if text in line:
            new_list.insert(idx, line)
            idx += 1
    file_read.close()
    lineLen = len(new_list)
    for i in range(lineLen):
        dateSubject, WorkersList= extractExcel(excel_path=values["-IN-"])  
        subject= new_list[i].replace("#{ENTER SUBJECT AFTER THIS} ","") #replaces str "{ENTER SUBJECT HERE"
        subject= subject.replace("{DATESUBJECT}", dateSubject)
    return subject

# Sends email to everyone in the Excel file. (BCC, will not show who else receieved an email). # Throws error if wrong Excel chosen. 
def Send_Email():
    noComments= cleanedEmailText() # Safety measure if user does not press "Preview Email"
    errCode= extractExcel(excel_path=values["-IN-"])
    subject= readSubjectName()
    if errCode != 1:
        dateSubject, WorkersList= extractExcel(excel_path=values["-IN-"])  
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.BCC = ";".join(WorkersList)
        mail.Subject = subject
        mail.Body= noComments
        mail.Send()
        sg.popup_no_titlebar("Email sent!", auto_close=True, auto_close_duration= 1.1, icon=Logo)
    else:
        sg.popup_error("Please choose another Excel File!", icon=Logo)
        # return("Please choose another Excel File!")

# View Excel File code # New Excel Preview - Used Tkinter
def viewExcel(excel_path):
    df= pd.read_excel(excel_path)
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name
    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview
    return None
def clear_data():
    tv1.delete(*tv1.get_children())
    return None

# Throws errors if filepath is not chosen
def validPath(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Please select a file path", title="Error", modal=True, icon=Logo)
    return False
    
# Resource Path - used for emailBody.txt (creates a temp file - removed per user request )
# Basically, would create a default emailBody.txt that could be edited, but WOULD be reverted to default emailBody.txt on next app use.
def resource_path(relative_path): #using for Logo now
    try:
        base_path= sys._MEIPASS
    except Exception:
        base_path= os.environ.get("_MEIPASS2", os.path.abspath("."))
    return os.path.join(base_path, relative_path)
Logo=resource_path(icon0)

# "Edit Email" pop up window.
def editMailPopup(text):
    multiline = sg.Multiline(text, size=(80, 20), reroute_cprint=True, key="-TEXT-") #key="ll")
    layout = [[multiline], [sg.Button('Save', size= (4,2)), sg.Button('Exit', size= (4,2))]]



    window = sg.Window('Edit Email Body', layout, modal=True, element_justification='r', icon=Logo)
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == "Save": 
            with open(("emailBody.txt"), "w" ) as file: #removed resource_path in with open 
                file.write(values["-TEXT-"])
            sg.popup("Saved!", auto_close=True, auto_close_duration= 1.1, icon=Logo)
    window.close()

# "Preview Email" pop up window. Had to create own func because previous one would save the .txt if button selected in "Preview" mode.
def previewMailPopup(text):
    multiline = sg.Multiline(text, size=(80, 20), reroute_cprint=True, key="-TEXT-") #key="ll")
    layout = [[multiline], [sg.Button('Exit', size= (4,2))]]
    window = sg.Window('Preview Email', layout, modal=True, element_justification='r', icon=Logo) 
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
    window.close()

# Shows "Cleaned Email Body." emailBody.txt has comments in the file, this code disregards the comments and emails the rest of msg. # Throws error if wrong Excel chosen. 
def cleanedEmailText():
    errCode= extractExcel(excel_path=values["-IN-"])
    if errCode != 1:
        dateSubject, WorkersList = extractExcel(excel_path=values["-IN-"])
        with open(("emailBody.txt"), "r") as file: #removed resource_path in with open 
            data = file.read()
            noComments= re.sub(r'(?m)^ *#.*\n?', '', data)
            noComments= noComments.replace("{DATESUBJECT}", dateSubject)
            noComments= noComments.replace("{SIGNATURE}", signature)
        return(noComments)
    else:
        # sg.popup_error("Please choose another Excel File!")
        return("Please choose another Excel File!")
        
# GUI
sg.theme("BlueMono")
layout= [[sg.Text("Input Excel File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),), s=12)], 
    [sg.Exit(s=10), sg.Button("Edit Email Body", s=12), sg.Button("Preview Email", s=12), sg.Button("View Excel File", s=12), sg.Button("Send Email", s=12)],]

window= sg.Window("Autosend Email", layout, icon=Logo)

while True:
    event, values= window.read()
    print(event, values)
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break

    # Replaces the Email Body text with var DATESUBJECT. 
    if event == 'Preview Email':
        if validPath(values["-IN-"]):
            noComments= cleanedEmailText()
            previewMailPopup(noComments)
        
    # Edit Email Body
    if event == "Edit Email Body":
        with open(("emailBody.txt"), "r") as file: #removed resource_path in with open 
            data = file.read()
        editMailPopup(data)

    if event == "View Excel File":
        if validPath(values["-IN-"]):           # Error message if Path not selected 
            # Start of Tkinter Code
            root = tk.Tk()
            root.title("Excel File Preview")
            root.geometry("500x520") 
            root.iconbitmap(Logo)
            root.pack_propagate(False)
            root.resizable(0, 0) #Window fixed in size
            style = ttk.Style(root)
            style.theme_use("clam")
            excelFrame = tk.LabelFrame(root) # Frame for TreeView (Excel Preview)
            excelFrame.place(height=450, width=500)
            file_frame = tk.Label(root) # Frame for load button
            file_frame.place(height=36, width=93, rely=0.88, relx=0.77)
            ExitButton = tk.Button(file_frame, text="Exit", height=2, width=10, compound="c", command=root.destroy)
            ExitButton.place(rely=0, relx=0)
            tv1 = ttk.Treeview(excelFrame)  ## Treeview Widget
            tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).
            # Scroll bar settings
            treescrolly = tk.Scrollbar(excelFrame, orient="vertical", command=tv1.yview) #update the Y-axis view of the widget
            treescrollx = tk.Scrollbar(excelFrame, orient="horizontal", command=tv1.xview) #update the X-axis view of the widget
            tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # Add scrollbars to treeview 
            treescrollx.pack(side="bottom", fill="x") #scrollbar fill the X-axis
            treescrolly.pack(side="right", fill="y") #scrollbar fill the Y-axis
            # End of Tkinter Code

            viewExcel(values["-IN-"])
    if event == "Send Email":
        if validPath(values["-IN-"]):           # Error message if Path not selected 
            Send_Email()

# root.mainloop() #not needed
window.close()

# Use this code in the terminal to create executable app 
# pyinstaller --onefile --noconsole --add-data emailBody.txt;. TestGUI.py
# or this with custom iso
# pyinstaller --onefile --noconsole --add-data emailBody.txt;. --add-data favicon.ico;. --icon favicon.ico TestGUI.py
