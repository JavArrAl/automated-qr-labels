import tempfile
import re
import os.path
from datetime import datetime
import pandas as pd
import numpy as np
from docx import Document
from docxtpl import DocxTemplate
import qrcode

import variableFile

'''
Code that contains the main backend classes and function of the LabelCode
This code needs at least two files to work properly.
An excel file populated with the data corresponding to the pumps
And a docx file corresponding to the template to be populated.
All relevant values are obtained from the Template file. Including:
- Number of labels
- Columns to select from the excel
'''

class WrongXlFile(Exception):
    pass


class WrongDocxFile(Exception):
    pass


class MissingXlFile(Exception):
    pass

class EmptyTemplate(Exception):
    pass


class EmbeddedFileError(Exception):
    pass 


class XlFile:
    def __init__(self,pathFile):
        self.pathFile = pathFile # File path of the docx template
        self.filt = False
        self.xlFilt = ''
        self.xlFiltVals = []
        self.readFile()
  
    def readFile(self):
        try:
            if self.pathFile.split('.')[-1] == 'xls': # Case not handled by openpyxl
                self.xlData = pd.read_excel(self.pathFile)
            else:
                self.xlData = pd.read_excel(self.pathFile,engine='openpyxl')
        except:
            raise WrongXlFile
    
    def selectColumns(self,xlParms =None):
        '''
        Receives the columns needed and returns those unique values.
        If filter is active, receives two argumens 'column' and 'value' from calling class
        Defined by user or 'Equimpent model' and 'BODYGUARD 323' by default
        Columns with Timestamp are filtered so they contain date only, formated as DD/MM/YY
        '''
        if not self.filt:
            xlData = self.xlData.copy()
        elif self.filt:
            xlData = self.xlData[self.xlData[self.xlFilt].isin(self.xlFiltVals)].copy()

        for column in xlData.columns:
            if xlData[column].dtype == '<M8[ns]':
                xlData[column] = pd.DatetimeIndex(xlData[column], normalize = True).strftime('%d-%m-%Y')
        
        if xlParms:
            return xlData[xlParms].dropna(subset = xlParms)
        elif not xlParms:
            return xlData
    
    def returnColumns(self):
        return list(self.xlData.columns)
    
    def returnValues(self,column):
        #return list(set(self.xlData[column].dropna()))
        return self.xlData[column].dropna().unique()

    def setFilter(self,column,values):
        '''Sets filter parameters.
        This shouldn't be here. This should be placed on the DocxFile class
        However, filter linked to excel on interface.
        Easier to keep it here and duplicate values
        ''' 
        self.xlFilt = column
        self.xlFiltVals = values
    

class DocxFile:
    def __init__(self,pathFile,xlClass):
        self.nameFile = '' # Name file wich could correspond to the template copied files
        self.pathFile = pathFile # File path of the docx template
        self.paramTmp = [] # Parameters required to populate the template.
        self.filt = False # Bolean to activate filter
        self.xlFilt = '' # Column excel filter
        self.xlFilVals = [] # Value to filter excel column selected
        self.filtIndx = [] # Index of selected rows (with or without filter). Was used for QR images. Not in use atm
        self.xlClass = xlClass # XlClass which handles unique excel file
        self.dictKeys = []
        self.context = [] # Dictionary with values to populatein the templates
        self.listQR = []

        if not xlClass:
            raise MissingXlFile
        self.readDocx()

    def xlDataCaller(self):
        return self.xlClass.selectColumns(
            self.paramTmp)

    def readDocx(self):
        '''Function that reads the docx file
        if read for the first time, if collects all unique tags on template
        and number of total labels.
        Replacement of '_' is redundant but necessary to interact with excel.
        '''
        try:
            self.doc = DocxTemplate(self.pathFile)
        except:
            raise WrongDocxFile
        if not self.paramTmp:
            numRe = re.compile(r'\d+')
            strRe = re.compile(r'[a-z_A-Z]+')
            tagsSet = self.doc.get_undeclared_template_variables()
            try:
                self.numLbl = int(max(re.findall(numRe,str(tagsSet))))
            except:
                raise EmptyTemplate
            self.paramTmp = [i.replace('_',' ') for i in list(set(re.findall(strRe,str(tagsSet))))]
            self.paramTmp = sorted(self.paramTmp, key = self.xlClass.returnColumns().index) # Params allways sorted as in the excel

    def createDict(self):
        '''Creates dictionary combining the tags found on the template
        with its corresponding values from the excel(filtered)
        '''
        inx = 1
        for _ in range(len(self.xlDataCaller())):
            for param in self.paramTmp:
                self.dictKeys.append('{}{}'.format(param.replace(' ','_'),inx))
            inx += 1
            if inx > self.numLbl: inx = 1
        self.context = list(zip(self.dictKeys,self.xlDataCaller().to_numpy().flatten()))
    
    def labelGeneration(self,listQR,numTemp,localContext,nameTmp = 'Template'):
        '''Function that fills the template docxs
        '''
        self.readDocx()
        self.doc.render(localContext)
        i = 1
        for pic in listQR:
            try:
                self.doc.replace_pic('Dummy{}.png'.format(i),pic)
            except:
                raise EmbeddedFileError
            i += 1
        self.doc.save('{}{}{}{}.docx'.format(self.tempPath,nameTmp,numTemp,datetime.now().strftime("_%d%b%y_%H;%M")))
        
    def labelGenLauncher(self):
        '''Calls labelGeneration function as many times as needed 
        depening on the length of the excel.
        It takes in consideration the number of labels per docx page
        '''
        # Automatic QR folder creation and destruction.
        self.tempFoldQR = tempfile.TemporaryDirectory()
        self.pathPic = self.tempFoldQR.name + '\\'
        #self.createBarcode()
        self.createQR()
        self.createDict()
        try:
            os.mkdir(os.path.expanduser('~\\Desktop\\QR_Templates'))
            self.tempPath = os.path.expanduser('~\\Desktop\\QR_Templates\\')
        except FileExistsError:
            self.tempPath = os.path.expanduser('~\\Desktop\\QR_Templates\\')
        ini = 0
        for rows in range(0,len(self.xlDataCaller()),self.numLbl):
            fin = ini +(self.numLbl*len(self.paramTmp))
            self.labelGeneration(self.listQR[rows:rows+self.numLbl],
                rows,dict(self.context[ini:fin]))
            ini = fin
            #self.upgradePB(self.numLbl)

        # Necessary for a proper functioning
        self.tempFoldQR.cleanup()
        self.listQR = []

    def createBarcode(self):
        '''
        Deprecated. Better use createQR.
        Function that creates the QR codes
        '''
        for index,row in self.xlDataCaller().iterrows():
            tempStr = ';'.join(map(str,row))
            qr = qrcode.QRCode(
                version = None, error_correction = qrcode.constants.ERROR_CORRECT_L,
                box_size=10, border=4)
            qr.add_data(tempStr)
            img = qr.make_image()
            self.listQR.append('{}QR{}'.format(self.pathPic,index))
            img.save('{}QR{}'.format(self.pathPic,index))   

    def createQR(self):
        '''Creates QR codes
        It looks the excel columns names on AI from variableFile
        If it exist, values will be used on QR code.
        If not, column will not be used in QR code (e.g: consumables)
        It matches AI with its corresponding value from the excel and
        generates a single string which will be converted to QR.
        If values are NaN (e.g. no docking station) they are not included in the QR
        However, the function will create a QR code if any field within AI's has a value.
        This bug is not fixed at the moment.
        '''

        tempVals = self.xlClass.selectColumns() # All values xl with formated Timestamps
        tempListAI = [None]*len(tempVals.columns)
        count = 0
        for col in tempVals.columns: # Create list AI corresponding Excel columns. Checks xl colums with AI dict of possible coincidences
            for key in variableFile.AI.keys():
                if col in variableFile.AI[key]:
                    tempListAI[count] = key
            count += 1
        
        tempTupList = []
        for lt in tempVals.to_numpy(): # creates tuples(AI,Value)
            tempTupList.append(list(zip(tempListAI,lt)))        
        
        countItem = 0
        for item in tempTupList:
            qr = qrcode.QRCode(
                version = None, error_correction = qrcode.constants.ERROR_CORRECT_L,
                box_size=10, border=4)
            temp = dict(item)
            temp.pop(None) # Discard non existent AI (excel column not in use for QR)
            new_temp = {ky:va for ky,va in temp.items() if pd.notna(va)} # Create dictionary without Nan values (e.g: No DS)
            new_temp = list(new_temp.items())
            qr.add_data(''.join([''.join(map(str,i)) for i in new_temp])) # Joins all tuples in str, then all those str in a single str
            img = qr.make_image()
            self.listQR.append('{}QR{}'.format(self.pathPic,countItem))
            img.save('{}QR{}'.format(self.pathPic,countItem))
            countItem += 1
            
    # Functions to update progressbar.
    def savePB(self,frame):
        self.progBar = frame
    
    def upgradePB(self,step):
        self.progBar.prgBar['value'] += step


if __name__ == "__main__":
    

    xlFile = "C:/Users/Javier/Documents/Projects/Docx Labels/Originals/" + "tmpBD1B.xls"
    lblDoc = "C:/Users/Javier/Documents/Projects/Docx Labels/Originals/" + "LabelsJavi2.docx"

    excel = XlFile(xlFile)
    docx = DocxFile(lblDoc,excel)
    docx.labelGenLauncher()
    # xlParmsLIst(xlParms,lbs)
    # xlData = readFile(xlFile, xlParms)
    # print(xlData)
    # for rows in range(0,len(xlData),lbs):
    #     nameImg = createBarcode(xlData[rows:rows+lbs],picPath)
    #     labelsWord(lblDoc,xlData[rows:rows+lbs],nameImg,rows,tempPath)
