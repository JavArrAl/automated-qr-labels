import tkinter as tk


# I know... it is a weak pass. It is not meant to be unhackeable.
# The excel sheet is locked to avoid possible unwanted typos
TEMPLATE_PASS = 12345

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

PUMPS_MODELS = {
    'AMBIX ACTIVE KIT': ['AMBIX ACTIVE KIT', 'AMBIX ACTIVE', 'AMBIX ACTIV', 'AMBIX', 'Ambix active', 'ambix active'],
    'BODYGUARD 323 KIT':['BODYGUARD 323 KIT', 'BG323', 'BG 323', 'BODYGUARD 323', 'BODYGUARD323'],
    'BODYGUARD CV323': ['BODYGUARD CV323', 'BG323CV', ' BG 323 CV', 'BG323 CV', 'BG CV 323', 'BGCV323','BODYGUARD 323 CV', 'BODYGUARD323CV', 'BODYGUARD CV 323', 'BODYGUARDCV 323'],
    'BODYGUARD DUO': ['BODYGUARD DUO', 'BG121CV', 'BGCV121','BG CV 121', 'BG 121 CV', 'BODYGUARD 121 CV', 'BODYGUARD CV 121', 'BGDUO', 'BG DUO'],
    'CADD LEGACY PLUS PUMP': ['CADD LEGACY PLUS PUMP', 'CADD LEGACY', 'CADDLEGACY', 'LEGACY PLUS', 'CADD LEGACY PLUS'],
    'CADD PRISM': ['CADD PRISM', 'PRSIM'],
    'CADD SOLIS VIP': ['CADD SOLIS VIP', 'CADD SOLIS', 'SOLIS'],
    'CRONO 30ML PUMP': ['CRONO 30ML PUMP', 'CRONO 30', 'CRONO30'],
    'CRONO PCA PUMP': ['CRONO PCA PUMP', 'CRONO PCA', 'PCA'],
    'CRONO PID PUMP':['CRONO PID PUMP', 'CRONO PID', 'CRONO PID'],
    'CRONO SPID 100': ['CRONO SPID 100', 'SPID 100', 'CRONO 100', 'PID 100'],
    'CRONO SPID 50': ['CRONO SPID 50', 'SPID 50', 'CRONO 50', 'PID 50'],
    'BODYGUARD 121 KIT': ['BODYGUARD 121 KIT', 'BG121', 'BG 121', 'BODYGUARD 121', 'BODYGUARD121'],
    'DUAL SIGNATURE': ['DUAL SIGNATURE', 'SIGNATURE'],
    'INFUSOMAT SPACE PUMP': ['INFUSOMAT SPACE PUMP', 'INFUSOMAT SPACE'],
    'MPMLH+ SYRINGE DRIVER': ['MPMLH+ SYRINGE DRIVER', 'MP-MLH+', 'MPMLH+'],
    'RYTHMIC PN PLUS PUMP': ['RYTHMIC PN PLUS PUMP', 'RYTHMIC PN +', 'PN +', 'PN+', 'RYTHMIC PN+', 'RYTHMIC PN PLUS', 'MINI RYTHMIC PN+', 'MINI PN+'],
    'SAPPHIRE H100 PUMP': ['SAPPHIRE H100 PUMP', 'H100', 'SAPPHIRE H100'],
    'SECA 875': ['SECA 875', '875']
}

INSTRUCTIONS_SCAN = '1.- To create the handover form, first open/create/select the file.\n\
2.- Once opened, go to the file and start scanning QR codes.\n\
3.- To compare with the client request form, select the request excel file.\n\
Note: Excel will group devices by model and color duplicates in orange.'


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