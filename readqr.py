import win32com.client as win32
import pandas as pd
import numpy as np
import tkinter as tk

import variableFile

class WorkbookEvents:
    def OnSheetSelectionChange(self, *args):
        '''Possible use to determine where the user is
        Might be useful to check whether correcting of adding new lines
        '''
        pass
        
    def OnSheetChange(self, *args):
        variableFile.addressChanged = args[1].Address
        variableFile.changedValue.set(str(args[1].Value))
        

class XlReadWrite:
    '''Class that handles reading, cleaning, processing
    of the open excel. It checks if the open workbook is 
    the one designed as Handover form. Name should have been
    previously agreed.
    '''
    # TODO: if values are read from sheet. Message showing the name of sheet1 by default.
    def __init__(self,parentFrame,xlapp):
        self.xl = xlapp
        self.parent = parentFrame
        self.xlWorkbook = None
        #self.dfValues = pd.DataFrame()
        self.checkWb()
        self.readExcel()

    def checkWb(self):
        '''Check if workbook is the handover form
        '''
        # TODO: include msg if the name does not match
        # NOTE: This is considering that only the xlsx files is going to be opened
        # FIXME: This is not working properly
        try:
            # self.xl.Workbooks('readingx.xlsx')
            if not self.xlWorkbook:
                self.xl.Workbooks('readingx.xlsx')
                self.xlWorkbook = self.xl.Workbooks('readingx.xlsx')
                self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
                # self.readExcel()
            else:
                None
            self.parent.readyVar.set('Ready')
            self.parent.readLbl.config(foreground = 'green')
        except:
            self.parent.readyVar.set('Open handover form')
            self.parent.readLbl.config(foreground = 'red')
    
    def readExcel(self):
        '''Reads current excel and creates file with all pumps included
        File used to check duplicates and do analytics
        '''
        # TODO: Think what happens if all this is deleted after moving to a diferent notebook frame
        try:
            self.values = self.xlWorkbook.Worksheets('Sheet1').UsedRange.Value
        except:
            print('No sheet called Sheet1')
        self.heads = self.values[1] # Request form heads are in row 2
        vals = self.values[2:] # Values start at row 3
        self.xlHeadsAI = []
         # list with AIs corresponding to excel columns
         # removed first and last item corresponding to ()
         # FIXME: if the excel has a column that is not in AI variableFile it will be a problem
        for head in self.heads:
            for key in variableFile.AI.keys():
                if head in variableFile.AI[key]:
                    self.xlHeadsAI.append(key[1:-1]) 
        # Creates a df from the tuple of tuples with all the values in the excel
        tempMap = map(list,zip(*vals)) # Transposes tuples with excel values
        tempDict = dict(zip(self.heads,list(tempMap)))
        # REVIEW: Is this the optimal way? create the dict every time the user changes a value?
        self.dfValues = pd.DataFrame(tempDict)
        #self.dfValues.append(tempDict, ignore_index = True)

    def processChanges(self,n,m,x):
        '''Function that process changes on excel
        Splits last value changed by () to get AI and values
        Adds dictionary with AI as keys and values as values to existing df

        ''' 
        # global changedValue
        # global addressChanged
        # TODO: Think how to link last row in DataFrame with the Address of last modified value in excel
        # FIXME: If AI in QR but not in excel, new entry created in df with all NaN
        readQR = str(variableFile.changedValue.get())
        valsAI = [tuple(i.split(')')) for i in readQR.split('(')]
        tempList = []
        # creates a dictionary with QR AI and values
        # Appends dictionary to df
        # NOTE: headsAI dependent on Excel column names. If not correct, wrong results.
        # NOTE: If any value introdcued by hand has () it will be counted as QR read
        if valsAI[0][0] == '': # Input comes from QR
            for vals in valsAI:
                for head in self.xlHeadsAI:
                    if vals[0] == head:
                        tempList.append((self.heads[self.xlHeadsAI.index(head)],vals[1]))
                        break
            self.dfValues = self.dfValues.append(dict(tempList),ignore_index = True)
            self.writeExcel()
        else: # Update value introduced by user in dfValues
            self.readExcel()

    def writeExcel(self):
        ''' Function that writes last row introduced in the excel
        It considers that the firts two rows of the excel are:
            1.- Date of delivery
            2.- Column heads
        '''
        # FIXME: If values are deleted, those cells would be still counted as with values
        # FIXME: If user introduces value below last row, new QR readings will be written after it
        # TODO: Sorting algorithms. If new scanned pump corresponds to same group, place it below
        
        lastRow = self.dfValues.index[-1] + 3
        # NaN replaced with None for empty cells in excel
        newRow = list(self.dfValues.iloc[-1].replace({np.nan: None}))  
        iniCell = '$A${}'.format(lastRow)
        finCol = chr(len(self.heads) + 96).upper()
        finCell = '${}${}'.format(finCol,lastRow)
        cellRange = '{}:{}'.format(iniCell,finCell)
        # Delete last edited cell
        self.xlWorkbook.Worksheets('Sheet1').Range(variableFile.addressChanged).Value = None 
        self.xlWorkbook.Worksheets('Sheet1').Range(cellRange).Value = newRow
    

if __name__ == "__main__":
    root = tk.Tk()
    variableFile.init()
