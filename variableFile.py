import tkinter as tk



templatePass = 12345

AI = {
    '(01)': ['Job No', 'Job_No', 'Job Number', 'Job_Number'],
    '(02)': ['Job Status', 'Job_Status'],
    '(10)': ['Creation Date', 'Creation_Date'],
    '(11)': ['Call_Date', 'Call Date'],
    '(12)': ['Close_Date', 'Close Date'],
    '(13)': ['Work_End_Date', 'Work End Date','Service_Date', 'Service Date','SERVICE DATE','SERVICE_DATE'],
    '(21)': ['Serial_Number', ' Serial Number', 'S/N', 'S No', 'S_No', 'Serial No', 'Serial_No','S/No'],
    '(22)': ['DS Serial No', 'DS_Serial_No', ' Docking Station', 'Docking_Station', 'Docking Station SN', 'Docking_Station_SN', 'DS_SN', 'DS SN', 'DS SNo', 'DS_SNo', 'Docking STN', 'DOCKING STN'],
    '(30)': ['Equipment_Model', 'Equipment Model', 'Model', 'Device Type', 'Device_Type','MODEL', 'Pump type', 'PUMP TYPE', 'Pump Type','pump type'],
    '(90)': ['Settings', 'Configuration','SETTINGS'],
    '(91)': ['Consumables','CONSUMABLES']
}



def init():
    global changedValue
    global excelOpen
    changedValue = tk.StringVar()
    excelOpen = tk.BooleanVar()
    excelOpen.set(tk.FALSE)
    changedValue.set('')

global addressChanged
global previousValue

previousValue = None
addressChanged = None