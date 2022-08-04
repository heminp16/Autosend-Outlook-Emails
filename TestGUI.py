from importlib.resources import path
import PySimpleGUI as sg
import pandas as pd

# Main Code
def index_containing_substring(list, substring):
        for i, s in enumerate(list):
            if substring in s:
                return i
        return -1

def Send_Email(excel_path):
    df1= pd.read_excel(excel_path).dropna(how="all")  #Dropping entire rows with NA
    df1["Worker"]= df1["Worker"].astype("string") #Converting Worker name to Str
    WorkersList= df1["Worker Email"].dropna().to_list() #Extracting Worker Emails, and dropping NA (the Applied Filter row) 

    # Extracting Week Range from the "Applied Filter" cell block
    weekRange= df1.iloc[-1,0] # Selecting last row
    split= weekRange.split("\n") #Splitting the Applied Filter row to sections

    # Code chunk to figure out where the Week Range is (Index)
    substring= "Week Range is"
    indexLocation = index_containing_substring(split, substring) #prints out the index number of where "Week Range is" is located 

    rangeStr= split[indexLocation] # Selected the Week Range  #'Week Range is x/xx/20xx - x/xx/20xx'

    dateSubject= rangeStr.split()[3:] # Selecting only the Dates # ['x/xx/20xx', '-', 'x/xx/20xx']
    dateSubject= " ".join(dateSubject) # Concatenate previous list to create one str 'x/xx/20xx - x/xx/20xx'

    # Email Code -- Change to make customizable later

    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.BCC = ";".join(WorkersList)
    mail.Subject = "Missing Time Entry" + " "+ dateSubject
    mail.Body = ("""Hello,

    You are missing time entry for the week range of""" + " " + dateSubject + ". " "Please update those hours as soon as possible."
    """

    Thank You,
    Hemin""")
    mail.Send()
    sg.popup_no_titlebar("Email sent!")
        
# GUI
sg.theme("BlueMono")
layout= [[sg.Text("Input Excel File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),("CSV Files", "*.csv"),))], #see if works on windows, confirmed does not work on Mac
    [sg.Exit(), sg.Button("Send Email")],]

window= sg.Window("Autosend Email", layout)

while True:
    event, values= window.read()
    print(event, values)
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    if event == "Send Email":
        Send_Email(
            excel_path=values["-IN-"]
        )
        
window.close()