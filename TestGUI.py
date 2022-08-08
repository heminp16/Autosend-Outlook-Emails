from importlib.resources import path
from pathlib import Path
import os, sys, re
import PySimpleGUI as sg
import pandas as pd
import win32com.client as win32

# Main Code
def index_containing_substring(list, substring):
        for i, s in enumerate(list):
            if substring in s:
                return i
        return -1

# *************************IF wrong excel type, give error *************************
def extractExcel(excel_path):
    df= pd.read_excel(excel_path).dropna(how="all")  #Dropping entire rows with NA
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

def Send_Email():
    dateSubject, WorkersList= extractExcel(excel_path=values["-IN-"])
    # Email Code -- Change to make customizable later    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.BCC = ";".join(WorkersList)
    mail.Subject = "Missing Time Entry" + " "+ dateSubject
    mail.Body= noComments
    mail.Send()
    sg.popup_no_titlebar("Email sent!")
    
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


def resource_path(relative_path):
    try:
        base_path= sys._MEIPASS
    except Exception:
        base_path= os.environ.get("_MEIPASS2", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

# EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING
# # Del or use the code I found  
def popup(text):
    multiline = sg.Multiline(text, size=(80, 20), reroute_cprint=True, key="ll")
    layout = [[multiline], [sg.Button('Save')], [sg.Button('Exit')]]
    window = sg.Window('Title', layout, modal=True)

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == "Save":
            with open(resource_path("emailBody.txt"), "r") as file: 
                window['Save'].update(value=file)
    window.close()

def cleanedEmailText():
    dateSubject, WorkersList = extractExcel(excel_path=values["-IN-"])
    with open(resource_path("emailBody.txt"), "r") as file:
        data = file.read()
        noComments= re.sub(r'(?m)^ *#.*\n?', '', data)
        noComments= noComments.replace("{DATESUBJECT}", dateSubject)
    return(noComments)

# def settingWindow(settings):
#     sg.popup_scrolled(settings, title= "Current Settings")

# EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING

# def mainWindow():
    # GUI
sg.theme("BlueMono")
layout= [[sg.Text("Input Excel File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),("CSV Files", "*.csv"),))], #see if works on windows, confirmed does not work on Mac
    [sg.Exit(), sg.Button("Edit Email Body"), sg.Button("View Excel File"), sg.Button("Send Email")],]

window= sg.Window("Autosend Email", layout)

while True:
    event, values= window.read()
    print(event, values)
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    #
    #
    # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING
    # Del or use the code I found 
    if event == 'Edit Email Body':
        if validPath(values["-IN-"]):    
            noComments= cleanedEmailText()
            popup(noComments)
        else:
            with open(resource_path("emailBody.txt"), "r") as file:
                data = file.read()
            popup(data)

        
    # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING # EDITING
    # 
    #     
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
