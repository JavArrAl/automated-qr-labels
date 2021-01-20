import re
import win32com.client as win32
import datetime
import os.path
import shutil

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
    
    def OnBeforeClose(self, *args):
        '''Event before the workbook closes and before asking
        to save changes.
        Possible use to trigger format before closing
        '''
        #if not win32.GetActiveObject('Excel.Application'): # Checks if excel is still running. 
        variableFile.excelOpen.set(tk.FALSE)
        pass
        

class XlReadWrite:
    '''Class that handles reading, cleaning, processing
    of the open excel. It checks if the open workbook is 
    the one designed as Handover form. Name should have been
    previously agreed.
    WorkFlow:
        - Workbook open?
            - Yes: Do name matches pattern?
                - Yes: Use that file
                - No: Go to open file
            - No: Go to folder, any name with temporal name?
                - Yes: Open that file
                - No: Check dates of files. Any one with an incoming date?
                    - Yes: Open that file
                    - No: Create a new file with temporal name

    After the workbook is opened, if closed:
        - Call checkExistXl. New app has been open?
            - Yes: checkExistWorkbook. Workbook exists?
                - Yes: Use that file
                - No: Keep waiting until new file is opened
    '''
    # TODO: bring selected workbook to the top
    # TODO: when a file is closed, let the user know the file is no longer open
    def __init__(self,parentFrame):
        self.xl = None
        self.parent = parentFrame
        self.xlWorkbook = None
        self.dirPath = os.path.expanduser('~\\Desktop\\REQUEST FORMS')
        #self.dfValues = pd.DataFrame()
        # self.checkWb()
        # self.readExcel()
    
    def openXl(self):
        '''Attempt to open excel
        If not  open, launches excel.
        '''
        self.restartObjects()
        try:
            self.xl = win32.GetActiveObject('Excel.Application')
        except:
            try:
                self.xl = win32.Dispatch('Excel.Application')
            except:
                self.parent.readyVar.set('Excel not available. Please make sure excel is installed')# Gets name of file
                self.parent.readLbl.config(foreground = 'red')
    
    def restartObjects(self):
        '''Sets all win32 objects references to None
        This is redundant, but was necessary to check possible problems
        '''
        self.xl = None
        self.xlWorkbook = None
        self.xlWorkbookEvents = None

    def checkNameDate(self,wbNames):
        '''Checks name file matches desired name
        Returns name of latest delivery date
        If name is "TEMPORAL REQUEST FORM" that file shoudl be used
        '''
        lastDate = datetime.date(2000,1,1)
        lastName = None
        for wbName in wbNames:
            tempName = re.search('TEMPORAL REQUEST FORM',wbName)
            name = re.search('REQUEST FORM',wbName)
            date = re.search(r'\d{2}.\d{2}.\d{2}',wbName)
            if tempName:
                return wbName, None
            if name and date:
                try:
                    curDate = datetime.datetime.strptime(date.group(0),'%d.%m.%y').date()
                    if curDate > lastDate: 
                        lastDate = curDate
                        lastName = wbName
                except: # Possibly the file has a date not valid, skip that file
                    continue
        return lastName, lastDate
    
    def checkDate(self,date):
        '''Function that checks the date is valid
        '''

        dateRegx = re.compile(r'(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|(?:(?:29|30)(\/|-|\.)(?:0?[13-9]|1[0-2])\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0?2\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)(?:(?:0?[1-9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})')
        corrDate = re.search(dateRegx, date)
        if corrDate:
            return date
        else:
            raise ValueError 

    def openWb(self,filePath):
        '''Opens an excel file selected by the user
        That file will be the one to work with
        '''

        self.openXl()    
        try:
            self.xlWorkbook = self.xl.Workbooks.Open(filePath)
            self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
            self.parent.readyVar.set('{}'.format(filePath.split('/')[-1]))# Gets name of file
            self.parent.readLbl.config(foreground = 'green')
            self.xl.Visible = True
            # variableFile.changedValue.trace('w',self.processChanges)
            self.readExcel()            
        except:
            self.parent.readyVar.set('ERROR opening excel. Contact support')       
        
    
    def fillXlOpenList(self):
        '''Creates a list with all opened excel files
        This list corresponds to the values of combobox on GUI
        '''

        try:
            self.xl = win32.GetActiveObject('Excel.Application')
            if self.xl.Workbooks.Count == 0:
                return []
            else:
                xlList = []
                for xl in range(1,self.xl.Workbooks.Count + 1):
                    xlList.append(self.xl.Workbooks(xl).Name)
                return xlList
        except:
            return []
        
    def selectWbActive(self,name):
        '''Sets the selected workbook as working workbook
        '''
        self.openXl()
        try:
            self.xlWorkbook = self.xl.Workbooks(name)
            self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
            self.parent.readyVar.set(name)
            self.parent.readLbl.config(foreground = 'green')
            self.xl.Visible = True
            self.readExcel()
        except:
            self.parent.readyVar.set('Error with excel file. Please panic')
            self.parent.readLbl.config(foreground = 'red')

    
    def newWb(self,date=None):
        '''Creates new request file in REQUEST FORM folder
        If folder does not exists, it is created.
        The new file is a copy of the template included with the program.
        The file is named as REQUEST FORM + DATE.
        The date is introduced by the user through GUI
        '''
        # try:
        #     self.xl = win32.GetActiveObject('Excel.Application')
        # except:
        #     self.xl = win32.Dispatch('Excel.Application')
        self.openXl()

        try:
            corrDate = self.checkDate(date)
            name = 'REQUEST FORM {}.xlsx'.format(corrDate)

            if not os.path.isdir(self.dirPath):
                os.mkdir(os.path.expanduser(self.dirPath))
            source = os.path.join(os.path.dirname(__file__),'templates','REQUEST FORM TEMPLATE.xlsx')
            destiny = os.path.join(self.dirPath,name) 
            try:
                shutil.copy(source, destiny)
                self.xlWorkbook = self.xl.Workbooks.Open(destiny)
                self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
                self.xl.Visible = True
                self.parent.readyVar.set(name)
                self.parent.readLbl.config(foreground = 'green')
                # variableFile.changedValue.trace('w',self.processChanges)
                self.readExcel()
            except PermissionError:
                self.parent.fileExists()
        except ValueError:
            self.parent.wrognDate()

        
    
    # def LEGACYopenWb(self):
    #     '''Checks if REQUESTED FORMS folder exists in Desktop
    #     Reads files and tries to find one with an incoming date
    #     If not, new file is created with temporal name until delivery date selected
    #     '''
    #     if os.path.isdir(self.dirPath):
    #         filesFolder = os.listdir(self.dirPath)
    #         lastName,lastDate = self.checkNameDate(filesFolder)
    #         if lastDate:
    #             if lastDate > datetime.date.today():
    #                 self.xlWorkbook = self.xl.Workbooks.Open(os.path.join(self.dirPath,lastName))
    #                 self.xl.Visible = True
    #                 self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
    #                 self.parent.readyVar.set(lastName)
    #             else:
    #                 self.newWb()
    #         if not lastDate:
    #             self.xlWorkbook = self.xl.Workbooks.Open(os.path.join(self.dirPath,'TEMPORAL REQUEST FORM.xlsx'))
    #             self.xl.Visible = True
    #             self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
    #             self.parent.readyVar.set('TEMPORAL REQUEST FORM.xlsx')
    #     else:
    #         self.newWb()
    
    # def LEGACYcheckNewWb(self):
    #     '''If a workbook is opened it is used as workbook.
    #     '''
    #     wbCount = self.xl.Workbooks.Count
    #     if wbCount != 0:
    #         self.xlWorkbook = self.xl.Workbooks(self.xl.Workbooks(1).Name)
    #         self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
    #         self.parent.readyVar.set(self.xl.Workbooks(1).Name)
    #     else:
    #         raise ValueError

    # def LEGACYcheckWb(self):
    #     '''Check if workbooks are open, checks names
    #     matches with template and finds coming dates.
    #     If not workbooks opened searches for files in folders.
    #     If file with coming date exists, it is used, if not
    #     a new file is created with TEMPORAL REQUEST FORM.
    #     This file will be used whenever it exists and should be
    #     always renamed as soon as possible.
    #     '''
    #     # try:
    #     #     self.xl.Workbooks('REQUEST FORM TEMPLATE.xlsx')
    #     #     if not self.xlWorkbook:
    #     #         #self.xl.Workbooks('REQUEST FORM TEMPLATE.xlsx')
    #     #         self.xlWorkbook = self.xl.Workbooks('REQUEST FORM TEMPLATE.xlsx')
    #     #         self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
    #     #         # self.readExcel()
    #     #     else:
    #     #         None
    #     #     self.parent.readyVar.set('Ready')
    #     #     self.parent.readLbl.config(foreground = 'green')
    #     # except:
    #     #     self.parent.readyVar.set('Open handover form')
    #     #     self.parent.readLbl.config(foreground = 'red')

    #     wbCount = self.xl.Workbooks.Count
    #     if wbCount == 0:
    #         self.openWb()
    #     else:
    #         # If any workbook exists, check its names to find file
    #         # If any matches, checks date. If multiple matches, opens most recent
    #         # If none matches opens a new workbook
    #         wbNames = [self.xl.Workbooks(i).Name for i in range(1,wbCount)]
    #         lastName,_ = self.checkNameDate(wbNames)
    #         if lastName:
    #             self.xlWorkbook = self.xl.Workbooks(lastName)
    #             self.xlWorkbookEvents = win32.WithEvents(self.xlWorkbook,WorkbookEvents)
    #             self.parent.readyVar.set(lastName)
    #         elif not lastName:
    #             self.openWb()

# TODO: re-think this whole part.
    def readExcel(self):
        '''Reads current excel and creates a file with all pumps included
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
         # FIXME: if the excel has a column that is not in AI variableFile it will be a problem
        for head in self.heads:
            for key in variableFile.AI.keys():
                if head in variableFile.AI[key]:
                    self.xlHeadsAI.append(key[1:-1]) # removed first and last item corresponding to ()
        # Creates a df from the tuple of tuples with all the values in the excel
        tempMap = map(list,zip(*vals)) # Transposes tuples with excel values
        tempDict = dict(zip(self.heads,list(tempMap)))
        # REVIEW: Is this the optimal way? create the dict every time the user changes a value?
        self.dfValues = pd.DataFrame(tempDict)
        self.dfValues.dropna(how = 'all', inplace = True) # Deletes included NaN rows
        print('From readExcel')
        print(self.dfValues)
        #self.dfValues.append(tempDict, ignore_index = True)

    def processChanges(self,n,m,x):
        '''Function that process changes on excel
        Splits last value changed by () to get AI and values
        Adds dictionary with AI as keys and values as values to existing df
        ''' 
        # TODO: Implement discard changes if read from QR on existing cell
        readQR = str(variableFile.changedValue.get())
        print(readQR)
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
            if tempList: # Only append values if list not empty            
                self.dfValues = self.dfValues.append(dict(tempList),ignore_index = True)
                self.dfValues.replace({np.nan: None}, inplace = True)
                self.formatExcel() # formats the dataframe to oder devices by model
            self.writeExcel()
        else: # Update value introduced by user in dfValues
            # self.readExcel()
            modCell = variableFile.addressChanged.split('$')[1:]
            modCell[0] = ord(modCell[0].lower()) - 97 # Conver column letter to number
            try:
                self.dfValues.iloc[int(modCell[1]) - 2][modCell[0]] = readQR
            except IndexError: # If user modifies cell after last row, read excel again
                # TODO: think what to do if user modifies row after last included value
                pass
        # TODO: after processing changes, trigger the table update (if exists)
    

    def writeExcel(self):
        ''' Function that writes last row introduced in the excel
        It considers that the firts two rows of the excel are:
            1.- Date of delivery
            2.- Column heads
        '''
        # FIXME: If values are deleted, those cells would be still counted as with values
        # FIXME: If user introduces value below last row, new QR readings will be written after it
        # TODO: Sorting algorithms. If new scanned pump corresponds to same group, place it below
        try:
            lastRow = self.dfValues.index[-1] + 3
        except IndexError:
            lastRow = 3
        # NaN replaced with None for empty cells in excel
        newRow = list(self.dfValues.iloc[-1].replace({np.nan: None}))  
        iniCell = '$A${}'.format(lastRow)
        finCol = chr(len(self.heads) + 96).upper()
        finCell = '${}${}'.format(finCol,lastRow)
        cellRange = '{}:{}'.format(iniCell,finCell)
        # Delete last edited cell
        self.xlWorkbook.Worksheets('Sheet1').Range(variableFile.addressChanged).Value = None 
        self.xlWorkbook.Worksheets('Sheet1').Range(cellRange).Value = newRow
        # self.formatExcel()
    
    def formatExcel(self):
        '''Function to groups all devices with same model together
        # NOTE: Should the paramenters to order the list be selected by the user?
        After reading QR, check model based on AI
        Sort the dataframe. Return new index last row introduced
        Insert new row on the corresponding excel position with (Range.Insert method)
        # LINK: https://docs.microsoft.com/en-us/office/vba/api/excel.range.insert
        '''
        self.dfValues.sort_values(by = 'MODEL', inplace = True, ignore_idex = True)
    
    def manageDuplicates(self):
        '''TODO: think what to do with the duplicates
        '''
        pass

class ClientRequest:
    '''Class that manages the reading and processing of the client requests
    Updates the table using updateTable method triggered by XlReadWrite
    '''
    def __init__(self,parent):
        self.myParent = parent # Frame class that invokes it
        self.file = self.myParent.filePathEntry

    def readExcel(self):
        ''' Function that reads excel and inserts into a df
        Currently it follows the Lloyds template (See Lloyd's template)
        '''
        self.clientXl = pd.read_excel(self.file, header = 1, usecols = 'B:H')

    
    def checkXL(self,xlClass):
        '''Checks if excel file is in use by the program.
        Updated if excel is opened. 
        '''
        pass
    
    def checkDate(self):
        '''If excel is open, checks delivery date
        Delivery date written in excel file if not present
        Changes the name of the excel file if it does not include date
        Pop up message if date excel file (name, cell) and client excel do not match.
        '''
        date = re.search(r'\d{2}-\d{2}-\d{4}',str(self.file))

    def createTable(self):
        '''Creates tkinter table with equipment from client excel
        First 3 columns (Model, Settings, Request) should be static
        Done column filled with pumps in excel file. Updated with updateTable
        '''
        # TODO: use Treeview with no children to create a table in tkinter
        # LINK: https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter/50651988#50651988

        pass

    def updateTable(self,xlClass):
        '''Updates values on "Done" column
        Triggered by the ReadXlClass when DataFrame is updated.
        '''
        # NOTE: TreeView set property to change value of item. 
        pass

    

if __name__ == "__main__":
    root = tk.Tk()
    variableFile.init()
