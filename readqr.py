import re
import win32com.client as win32
import datetime
import os.path
import shutil
from ast import literal_eval
import unicodedata

import pandas as pd
import numpy as np
import tkinter as tk


import variableFile

# (30)AMBIX ACTIV(21)21242770(13)22-12-2020(22)20468303
# (30)CRONO 30(21)NL0411.16(13)21-12-2020
#'$3:$3,$3:$24,$3:$25'

class WorkbookEvents:
    def OnSheetSelectionChange(self, *args):
        '''Possible use to determine where the user is
        Might be useful to check whether correcting of adding new lines
        '''
        variableFile.previousValue = args[1].Value
        
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
        # TODO: when excelOpen set to FALSE change the GUI label to gray when corresponds
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
    # TODO: when a file is closed, let the user know the file is no longer open
    def __init__(self,parentFrame):
        self.xl = None
        self.parent = parentFrame
        self.xlWorkbook = None
        self.dirPath = os.path.expanduser('~\\Desktop\\REQUEST FORMS')
    
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
            variableFile.excelOpen.set(tk.TRUE)
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
            variableFile.excelOpen.set(tk.TRUE)
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
                self.xlWorkbook.Worksheets('Sheet1').Range('$B$1').Value = corrDate
                variableFile.excelOpen.set(tk.TRUE) # NOTE: This could be done with events
                self.readExcel()
            except PermissionError:
                self.parent.fileExists()
        except ValueError:
            self.parent.wrognDate()

    def readExcel(self):
        '''Reads current excel and creates a file with all pumps included
        File used to check duplicates and do analytics
        '''
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
        tempDict = self.excelValToDict(vals)
        self.dfValues = pd.DataFrame(tempDict)
        self.dfValues = self.dfValues.convert_dtypes() # Converts columns types to the corresponding dtypes
        for col in self.dfValues.select_dtypes(include = 'string'):
            self.dfValues[col] = self.dfValues[col].str.normalize('NFKD') # Normalises unicode to include whitespaces (instead of \xa0)       

    def excelValToDict(self,vals):
        '''Converts tuples from excel into a dictionary
        Columns correspond to excel head values
        '''
        tempMap = map(list,zip(*vals)) # Transposes excel value tuples
        tempDict = dict(zip(self.heads,list(tempMap)))
        return tempDict

    def processChanges(self,n,m,x):
        '''Function that process changes on excel
        Splits last value changed by () to get AI and values
        Adds dictionary with AI as keys and values as values to existing df
        ''' 
        readQR = str(variableFile.changedValue.get())
        valsAI = [tuple(i.split(')')) for i in readQR.split('(')]
        tempList = []
        try: # If multiple cells selected consider deleting
            isDelete = all(item is None for tup in literal_eval(readQR) for item in tup)
        except:
            isDelete = False
        # creates a dictionary with QR AI and values
        # Appends dictionary to df
        # NOTE: headsAI dependent on Excel column names. If not correct, wrong results.
        if valsAI[0][0] == '' and not isDelete: # Input comes from QR
            for vals in valsAI:
                for head in self.xlHeadsAI:
                    if vals[0] == head:
                        tempList.append((self.heads[self.xlHeadsAI.index(head)],vals[1]))
                        break
            if tempList: # Only append values if list not empty            
                self.dfValues = self.dfValues.append(dict(tempList),ignore_index = True)
                self.dfValues.replace({np.nan: None}, inplace = True)               
                # self.formatExcel() # formats the dataframe to oder devices by model
            self.writeExcel()
        elif isDelete:
            self.deleteCell(literal_eval(readQR))
            self.formatExcel()
        else: # Update value introduced by user in dfValues
            modCell = self.multipleCellChange()
            try:
                self.dfValues.iloc[modCell[0],modCell[1]] = readQR
            except IndexError: # If user modifies cell after last row, read excel again
                self.readExcel()
    # If client request files has been loaded, update the table
        if self.parent.myParent.existsTable():
            self.parent.myParent.updateTable(self.returnCountDevices())
    
    def deleteCell(self,read):
        '''Function that proccesses the deleting of a cell/Range of cells
        '''
        modCell = self.multipleCellChange()
        tempDict = self.excelValToDict(read)
        tempDf = pd.DataFrame(tempDict)
        try:
            self.dfValues.iloc[modCell[0],modCell[1]] = tempDf.values
        except ValueError:
            modCell = self.returnOverRange()   
            self.dfValues.iloc[modCell[0],modCell[1]] = None
        
    def returnOverRange(self):
        '''If selected range to delete is larger than the actual dataframe
        Start at the first cell and complete from there
        '''
        singleCell = variableFile.addressChanged.split(':')
        firstCell = self.multipleCellChange(singleCell = singleCell[0])
        lastCell = self.multipleCellChange(singleCell = singleCell[1])
        if lastCell[0] > self.dfValues.shape[0]:
            lastCell[0] = self.dfValues.shape[0]
        if lastCell[1] > self.dfValues.shape[1]:
            lastCell[1] = self.dfValues.shape[1]
        return [slice(firstCell[0],lastCell[0]+1), slice(firstCell[1],lastCell[1]+1)]

    def multipleCellChange(self, singleCell = None):
        '''Returns the range of cell/s used
        When using slice objects, adding one to include the last value
        singleCell used when only one Cell is to be translated into df indexes
        '''
        if singleCell == None:
            address = variableFile.addressChanged
        else:
            address = singleCell
        if ':' in address:
            cells = [c for a in address.split(':') for c in a.split('$') ]
            cellRows = [int(cells[2]) - 3 , int(cells[5]) - 3]
            cellCols = [ord(cells[1].lower()) - 97, ord(cells[4].lower()) - 97]
            if cellRows[0] == cellRows[1]: # Only one row selected
                if cellCols[0] == cellRows[1]:
                    return [cellRows[0], cellCols[0]]
                else:
                    return [cellRows[0], slice(cellCols[0],cellCols[1]+1)]
            elif cellRows[0] == cellRows[1]: # Only one column selected
                return [slice(cellRows[0],cellRows[1]), cellCols[0] ]
            # If multiple rows and columns selected.
            return [slice(cellRows[0],cellRows[1]+1), slice(cellCols[0],cellCols[1]+1)]
        else:
            modCell = address.split('$')[1:]
            retCell = [int(modCell[1]) - 3, ord(modCell[0].lower()) - 97] # Convert column letter to number
            return retCell    

    def writeExcel(self):
        ''' Function that writes last row introduced in the excel
        It assumes the firts two rows of the excel are:
            1.- Date of delivery
            2.- Column heads
        '''
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
        # Restore last edited cell
        # FIXME: The font and size of the previous value to be written on the cell has changed.
        # It could be related with the excel license, check with proper internet.
        self.xlWorkbook.Worksheets('Sheet1').Range(variableFile.addressChanged).Value = variableFile.previousValue
        self.xlWorkbook.Worksheets('Sheet1').Range(cellRange).Value = newRow
        self.formatExcel()
    
    def formatExcel(self):
        '''Function to groups all devices with same model together
        # NOTE: Should the paramenters to order the list be selected by the user?
        After reading QR, check model based on AI
        Sort the dataframe. Return new index last row introduced
        Insert new row on the corresponding excel position with (Range.Insert method)
        # LINK: https://docs.microsoft.com/en-us/office/vba/api/excel.range.insert
        # TODO: change Column A and C for parameters selected by the user. Multiple parameters could be used
        '''
        
        self.xlWorkbook.Worksheets('Sheet1').Unprotect(Password = variableFile.TEMPLATE_PASS)
        lastCol = chr(len(self.heads) + 96).upper()
        lastRow = self.dfValues.index[-1] + 3
        allCells = f'$A$3:${lastCol}${lastRow}' 

        self.removeEmptyRows(lastCol)

        # Excel sorting
        self.xlWorkbook.Worksheets('Sheet1').Range(allCells).Sort(
            Key1 = self.xlWorkbook.Worksheets('Sheet1').Range('$A$3'), # Model
            Key2 = self.xlWorkbook.Worksheets('Sheet1').Range('$C$3'), # Date
            Orientation = 1, #This should be 2, but seems to be wrong.
            DataOption2 = 1 # Treats text as numeric
        )

        self.readExcel() # "Update" the data frame

        self.manageDuplicates(lastCol)

        self.xlWorkbook.Worksheets('Sheet1').Protect(Password = variableFile.TEMPLATE_PASS)

    def removeEmptyRows(self,lastCol):
        '''Removes empty rows from excel and df
        '''
        wholeDfIndex = self.dfValues.index.to_list()
        nonNanDfIndex = self.dfValues.dropna(how = 'all').index.to_list()
        nanDfIndex = list(set(wholeDfIndex) - set(nonNanDfIndex))
        nanDfIndex.sort(reverse=True) # DataFrame indexes of whole NaN rows

        for item in nanDfIndex:
            row = f'$A${item + 3}:${lastCol}${item + 3}'
            self.xlWorkbook.Worksheets('Sheet1').Range(row).EntireRow.Delete(Shift = -4162)

        self.dfValues.dropna(how = 'all', inplace = True)

    def manageDuplicates(self, lastCol):
        '''After updating Df find duplicates in pandas
        Color cells with excel.
        '''
        duplicatesDevices = self.dfValues[self.dfValues.duplicated()].index.to_list()
        for item in duplicatesDevices:
            toColorRange = f'$A${item + 3}:${lastCol}${item + 3}'
            self.xlWorkbook.Worksheets('Sheet1').Range(toColorRange).Interior.ColorIndex = 44
    
    def returnCountDevices(self):
        '''Returns the current count of items on df by model
        Function used to update the table on GUI
        # NOTE: What if the pump is not in the pool of words?
        # NOTE: Time for ML?
        '''
        tempDf = self.dfValues.copy()
        tempDf['KEY'] = ''
        tempDf['COUNT'] = 0
        # tempDf.groupby(by = ['MODEL'])['COUNT'].count()
        tempDf.set_index(['MODEL'], drop = False, inplace = True)
        # Creates new column on tempDf with the corresponding key on clientDf
        for pump in tempDf['MODEL']:
            for key in variableFile.PUMPS_MODELS.keys():
                if pump in variableFile.PUMPS_MODELS[key]:
                    tempDf.loc[pump,'KEY'] = key
        
        # Returns pandas series with the key count
        return tempDf.groupby(by=['KEY'])['COUNT'].count().convert_dtypes()
        

class ClientRequest:
    '''Class that manages the reading and processing of the client requests
    Updates the table using updateTable method triggered by XlReadWrite
    '''
    def __init__(self,parent):
        self.myParent = parent # Frame class that invokes it

    def readExcel(self):
        ''' Function that reads excel and inserts into a df
        Currently it follows the Lloyds template (See Lloyd's template)
        Returns df with two columns, pump type and its requests which are greater than 0
        '''
        self.file = self.myParent.filePathEntry
        self.clientXl = pd.read_excel(self.file, header = 1, usecols = 'B:H')
        # return self.clientXl[self.clientXl['Request'] != 0][['Pump Type','Request','Settings']]

        # If devices have different settings, the number in the table is total of same devices
        return self.clientXl[self.clientXl['Request'] > 0].groupby('Pump Type')['Request'].sum().reset_index()
    
    def checkDate(self):
        '''If excel is open, checks delivery date
        Delivery date written in excel file if not present
        Changes the name of the excel file if it does not include date
        Pop up message if date excel file (name, cell) and client excel do not match.
        '''
        date = re.search(r'\d{2}-\d{2}-\d{4}',str(self.file))

    # TODO: create methods to check pumps included
    # For example: 
    #       - If there are enough pumps of one type (be aware of duplicates): change color row table or similar
    #       - Include total of pumps scanned
    #       - Change color of last type of pump included

    

if __name__ == "__main__":
    root = tk.Tk()
    variableFile.init()
