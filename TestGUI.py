# from importlib.resources import path
from pathlib import Path #need
import os, sys, re
import PySimpleGUI as sg
import pandas as pd
# import win32com.client as win32

# Main Codes
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
        sg.popup_error("This might be the incorrect Excel file. Columns: 'Worker' or/and 'Worker Email' are missing")
        return errCode

# Sends email to everyone in the Excel file. (BCC, will not show who else receieved an email). # Throws error if wrong Excel chosen. 
def Send_Email():
    errCode= extractExcel(excel_path=values["-IN-"])
    if errCode != 1:
        dateSubject, WorkersList= extractExcel(excel_path=values["-IN-"])  
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.BCC = ";".join(WorkersList)
        mail.Subject = "Missing Time Entry" + " "+ dateSubject
        mail.Body= noComments
        mail.Send()
        sg.popup_no_titlebar("Email sent!")
    else:
        sg.popup_error("Please choose another Excel File!")
        # return("Please choose another Excel File!")

# View Excel File code
def viewExcel(excel_path):
    df= pd.read_excel(excel_path)
    filename= Path(excel_path).name
    sg.popup_scrolled( "=" * 50, df, title= filename)

# Throws errors if filepath is not chosen
def validPath(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Please select a file path")
    return False
    
# Resource Path - used for emailBody.txt
def resource_path(relative_path):
    try:
        base_path= sys._MEIPASS
    except Exception:
        base_path= os.environ.get("_MEIPASS2", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

# "Edit Email" pop up window. #Will edit file, but does not show edits until you exit. FIX THIS!
def editMailPopup(text):
    multiline = sg.Multiline(text, size=(80, 20), reroute_cprint=True, key="-TEXT-") #key="ll")
    layout = [[multiline], [sg.Button('Save')], [sg.Button('Exit')]]
    window = sg.Window('Title', layout, modal=True)
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == "Save": 
            with open(resource_path("emailBody.txt"), "w" ) as file:
                window["-TEXT-"].update(data)
                file.write(values["-TEXT-"])
    window.close()

# "Preview Email" pop up window. Had to create own func because previous one would save the .txt if button selected in "Preview" mode.
def previewMailPopup(text):
    multiline = sg.Multiline(text, size=(80, 20), reroute_cprint=True, key="-TEXT-") #key="ll")
    layout = [[multiline], [sg.Button('Exit')]]
    window = sg.Window('Title', layout, modal=True)
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
        with open(resource_path("emailBody.txt"), "r") as file:
            data = file.read()
            noComments= re.sub(r'(?m)^ *#.*\n?', '', data)
            noComments= noComments.replace("{DATESUBJECT}", dateSubject)
        return(noComments)
    else:
        # sg.popup_error("Please choose another Excel File!")
        return("Please choose another Excel File!")

# def settingWindow(settings):
#     sg.popup_scrolled(settings, title= "Current Settings")

# GUI
sg.theme("BlueMono")
layout= [[sg.Text("Input Excel File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),("CSV Files", "*.csv"),))], #see if works on windows, confirmed does not work on Mac
    [sg.Exit(), sg.Button("Edit Email Body"), sg.Button("Preview Email"), sg.Button("View Excel File"), sg.Button("Send Email")],]

window= sg.Window("Autosend Email", layout)

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
        with open(resource_path("emailBody.txt"), "r") as file:
            data = file.read()
        editMailPopup(data)

    if event == "View Excel File":
        if validPath(values["-IN-"]):           # Error message if Path not selected 
            viewExcel(values["-IN-"])
    if event == "Send Email":
        if validPath(values["-IN-"]):           # Error message if Path not selected 
            Send_Email()
        
window.close()

# if __name__ == "__main__":
#     settingPath= Path.cwd()
#     settings= sg.UserSettings(
#         path= settingPath, filename="config.ini", use_config_file=True
#     )
#     mainWindow() 


# pyinstaller TestGUI.py --onefile --add-data emailBody.txt;. --windowed
# pyinstaller  --onefile --windowed --add-data emailBody.txt TestGUI.py;.

# pyinstaller TestGUI.py --onefile --windowed ^ --add-data emailBody.txt;. 


# pyinstaller --onefile --noconsole --add-data emailBody.txt;included TestGUI.py --distpath .

#pyinstaller TestGUI.py --onefile --windowed ^ --add-data="emailBody.txt;." 


# WORKS YES
#pyinstaller --onefile --noconsole --add-data emailBody.txt;. TestGUI.py

############################################################################################
############################################################################################
################  Only thing left: Fix "Save" button when editing txt file  ################
############################################################################################
############################################################################################