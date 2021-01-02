import win32com.client as win32
import pandas as pd

import variableFile

class WorkbookEvents:
    def OnSheetSelectionChange(self, *args):
        '''Possible use to determine where the user is
        Might be useful to check whether correcting of adding new lines
        '''
        pass
        
    def OnSheetChange(self, *args):
        variableFile.changedValue.set(str(args[1].Value))
        variableFile.addressChanged = args[1].Address

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
        self.checkWb()

    def checkWb(self):
        '''Check if workbook is the handover form
        '''
        # TODO: include msg if the name does not match
        # NOTE: This is considering that only the xlsx files is going to be opened
        try:
            self.xl.Workbooks('readingx.xlsx')
            if not self.xlWorkbook:
                self.xlWorkbook = self.xl.Workbooks('readingx.xlsx')
                self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
                self.readExcel()
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
            self.values = self.xlWorkbook.Worksheets('Sheet1').UsedRange
        except:
            print('No sheet called Sheet1')
        # TODO: Find AI corresponding to heads
        self.heads = self.values[1] # Request form heads are in row 2
        vals = self.values[2:] # Values start at row 3
        self.xlHeadsAI = []
         # list with AIs corresponding to excel columns
         # removed first and last item corresponding to ()
         # FIXME: if the excel has a column that is not in variableFile this will create a problem
        for head in self.heads:
            for key in variableFile.AI.keys():
                if head in variableFile.AI[key]:
                    self.xlHeadsAI.append(key[1:-1]) 
        # Creates a df from the tuple of tuples with all the values in the excel
        tempMap = map(list,zip(*vals)) # Transposes tuples with excel values
        tempDict = dict(zip(self.heads,list(tempMap))) # Zips each column name with the list of values
        self.dfValues = pd.DataFrame(tempDict)  

    def processChanges(self):
        '''Function that process changes on excel
        ''' 
        # global changedValue
        # global addressChanged
        # TODO: Think how to link last row in DataFrame with the Address of last modified value in excel
        readQR = str(variableFile.changedValue)
        valsAI = [tuple(i.split(')')) for i in readQR.split('(')]
        tempList = []
        # creates a dictionary with QR AI and values
        # Appends dictionary to df
        # NOTE: headsAI dependent on Excel column names. If not correct, wrong results.
        if valsAI[0][0] == '': # Input comes from QR
            for vals in valsAI:
                for head in self.xlHeadsAI:
                    if vals[0] == head:
                        tempList.append(self.heads[self.xlHeadsAI.index(head)],vals[1])
                        break
            self.dfValues.append(dict(tempList),ignore_index = True)
        else:
            None


