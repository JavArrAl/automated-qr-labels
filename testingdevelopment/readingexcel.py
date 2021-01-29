import sys

import pandas as pd
import pythoncom
import win32com.client as win32
import tkinter as tk

class ApplicationEvents:
    def OnSheetActivate(self, *args):
        print('Something changed')


class WorkbookEvents:
    # def OnSheetSelectionChange(self, *args):


        # print(args)
        # print(args[1].Address)
        # args[0].Range('A1').Value = 'You selected cell' + str(args[1].Address)
        
    def OnSheetChange(self, *args):
        global changedValue
        changedValue.set(str(args[1].Value))

class MainFrame(tk.Frame):
    def __init__(self,parent):
        global changedValue
        self.Myparent = parent
        self.newVal = 0
        tk.Frame.__init__(self,parent)
        self.lbl = tk.Label(self, textvariable = changedValue)

        self.pack()
        self.lbl.pack(fill = 'both')

        self.xl = win32.GetActiveObject('Excel.Application')
        xlEvents = win32.WithEvents(self.xl, ApplicationEvents)
        self.xlWorkbook = self.xl.Workbooks('readingx.xlsx')
        xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)

        # changedValue.trace('w',update_lbl)

    def update_lbl(self,a,b,c):
        print('value changed')
        print(self.xlWorkbook.Worksheets('Sheet1').UsedRange) # Gathers all data in that worksheet
        


root = tk.Tk()
changedValue = tk.StringVar()

changedValue.set('')
frame = MainFrame(root)
changedValue.trace('w',frame.update_lbl)

root.mainloop()


# xl = win32.GetActiveObject('Excel.Application')

# xlEvents = win32.WithEvents(xl, ApplicationEvents)

# xlWorkbook = xl.Workbooks('readingx.xlsx',frame)

# xlWorkbookEvents = win32.WithEvents(xlWorkbook,WorkbookEvents)


# keepOpen = True

# while keepOpen:


#     #pythoncom.PumpMessages()

#     try:

#         if xl.Workbooks.Count != 0:
#             keepOpen = True

#         else:
#             keepOpen = False
#             xl = None
#             sys.exit()

#     except:

#         keepOpen = False
#         xl = None
#         sys.exit()

