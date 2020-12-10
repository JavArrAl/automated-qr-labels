from docx import Document
import pandas as pd
import numpy as np
import variableFile
import re
import os.path
import shutil
from docxtpl import DocxTemplate

import barcode
from barcode.writer import ImageWriter
import qrcode

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

## TODO: Class that handles xl files
# This functon includes all functions of this script that are not in docx
# inclu
class XlFile:
    def __init__(self,pathFile):
        self.pathFile = pathFile # File path of the docx template
        self.readFile()

        #This should be an automatic folder created to store de QR and
        # destroyed after the application is closed.
        # This should be probably located within collectionQRcode
        self.pathPic = 'C:/Users/Javier/Documents/Projects/Docx Labels/QRpng/'

    
    def readFile(self):
        try:
            self.xlData = pd.read_excel(self.pathFile)
            self.xlData = self.xlData[self.xlData.isnull().sum() == 0]
        except:
            raise WrongXlFile
    
    def selectColumns(self,xlParms,value = False):
    #Receives the columns needed and returns those unique values
        return self.xlData[xlParms]
    
    # Unnecesary
    # def selectFlatColumns(self,xlParams):
    #    return self.xlData[xlParams].to_numpy().flatten()

        
    def applyFilter(self,column,value = any):
        '''Function that applies filter to excel selected column
        based on paramteres employed by the user
        '''
        return self.xlData[xlData[column] == value]
    
    def createBarcode(self):
        '''
        Function that creates the QR codes after reading and cleaning the excel
        '''
        self.pathQR = []
        for index,row in self.xlData.iterrows():
            tempStr = ''
            ## REVIEW last column is a timestamp, it should be date
            ## REVIEW esto es super cutre, encontrar forma de hacerlo mas efectivo (maybe .tolist())
            for column in range(len(row)):
                    tempStr += str(row[column])
            if variableFile.CODE_TYPE == "Barcode": 
                barcode.get('code128',tempStr,writer=ImageWriter()).save('{}/BrcdPNG{}'.format(self.pathPic,row[1]))
            elif variableFile.CODE_TYPE == "QR":
                qr = qrcode.QRCode(version = None, error_correction = qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
                qr.add_data(tempStr)
                img = qr.make_image()
                self.pathQR.append('{}QR{}'.format(self.pathPic,index))
                img.save('{}QR{}'.format(self.pathPic,index))
        
    def collectionQRcode(self,selection):
        '''
        Function that returns the necessary list of QR image paths
        '''
        pass


## TODO: Class that works as a template docx handler
# - It will help manage all parameters related to the docx handler:
#   - Number of labels
#   - Parameters required (and excel file selection)
#   - Handle filter values
# - It will contain all the functions related with the template generation
# this will ease the calling of this functions without the necessity to pass multiple arguments
#   - labelGeneration
#   - labelGenLauncher

class DocxFile:
    def __init__(self,xlClass):
        self.numLbl = 0
        self.nameFile = '' # Name file wich could correspond to the template copied files
        self.pathFile = '' # File path of the docx template
        self.paramTmp = [] # Parameters required to populate the template.
        self.xlFilt = '' # Column excel selected
        self.xlFilVals = [] # Value to filter excel column selected
        self.xlClass = xlClass # XlClass which handles unique excel file
        self.dictKeys = []
        self.context = [] # Dictionary with values to look for in the templates
        
        # This should be created dinamicaly or selected by the user
        self.tempPath = "C:/Users/Javier/Documents/Projects/Docx Labels/finalTemplates/"

    def readDocx(self):
        '''Function that reads the docx file 
        '''
        try:
            self.doc = DocxTemplate(lblDoc)
        except:
            raise WrongDocxFile
        self.doc.render(self.context)
        # TODO: someway to calculate numLbl
        # TODO: someway to calculate paramTmp

    def createDict(self):
        for label in range(self.numLbl):
            for param in self.paramTmp:
                self.dictKeys.append('{}{}'.format(param.replace(' ','_'),label+1))
        self.context = dict(zip(self.dictKeys,\
            self.xlClass.selectColumns(self.paramTmp).to_numpy().flatten()))
    
    def labelGeneration(self,listQR,numTemp,nameTmp = 'Template'):
        '''Function that fills the template docxs
        '''
        self.doc.render(self.context)
        for pic in range(len(listQR)):
            self.doc.replace_pic('Dummy{}.png'.format(pic+1),self.listQR[pic])
        self.doc.save('{}{}{}.docx'.format(self.tempPath,nameTmp,numTemp))
        
    def labelGenLauncher(self):
        '''Calls labelGeneration function as many times as needed 
        depening on the length of the excel.
        It takes in consideration the number of labels per docx page
        '''
        ## TODO: think what to do with filtered things!!!!!!!!!
        self.createDict()
        self.listQR = self.xlClass.collectionQRcode(self.paramTmp,self.xlFilt)
        for rows in range(0,len(self.xlClass.selectColumns()),self.numLbl):
            self.labelGeneration(self.listQR[rows:rows+self.numLbl],rows)





def readFile(xlFile, xlParms):
    ''' Reads the excelp file and creates a data frame with
    the parameters selected by the user. It discards the N/A rows by default
    '''
    try:
        xlData = pd.read_excel(xlFile)
    except:
        raise WrongXlFile
    xlData = xlData[xlParms]
    for param in xlParms:
        xlData = xlData[xlData[param].notnull()]
    return xlData

## TODO: This is transfered to the DocxFile class
def auxLabelWord(auxlblDoc,xlData,picPath,tempPath,paramFilt = 'Equipment Model', model = 'BODYGUARD 323'):
    '''Invokes methods to create lables using only the models selected
    '''
    ## NOTE: This has not been tested.
    lbs = variableFile.NUM_LABELS_AUX # This should be obtained from the template
    xlData = xlData[xlData[paramFilt] == model] #This will be done by the filter function
    for rows in range(0,len(xlData),lbs):
        nameImg = createBarcode(xlData[rows:rows+lbs],picPath)
        labelsWord(auxlblDoc,xlData[rows:rows+lbs],nameImg,rows,tempPath,'Aux_Templates')
 

def labelsWord(lblDoc,xlData,nameImg,numTemp,tempPath,nameTmp = 'Templates'):
    ''' Creates the labels in a copy of the original template docx file.
    Creates a dictionary matching values from the excel introduced by the user
    to the tags in the docx template.
    '''
    try:
        doc = DocxTemplate(lblDoc)
    except:
        raise WrongDocxFile
    vals = xlData.to_numpy().flatten()
    context = dict(zip(variableFile.DICT_KEYS,vals))
    doc.render(context)
    for pic in range(len(nameImg)):
        doc.replace_pic('Dummy{}.png'.format(pic+1),nameImg[pic])
    doc.save('{}{}{}.docx'.format(tempPath,nameTmp,numTemp))

    

def createBarcode(xlData,picPath):
    ''' Creates the code. Can be QR or barcode depending on the user choice.
    This function stores generates pictures in a folder during the operation.
    Returns list with images names so they can be used by "labelsWord"
    '''
    ## NOTE: is it necesary to save the pictures? could they be directly included in the word?
    ## REVIEW at the moment it stores the barcodes as pictures in the computer. Check if this can be done in any other way. 

    nameImg = []
    for index,row in xlData.iterrows():
        tempStr = ''
        ## REVIEW last column is a timestamp, it should be date
        ## REVIEW esto es super cutre, encontrar forma de hacerlo mas efectivo (maybe .tolist())
        for column in range(len(row)):
                tempStr += str(row[column])
        if variableFile.CODE_TYPE == "Barcode": 
            barcode.get('code128',tempStr,writer=ImageWriter()).save('{}/BrcdPNG{}'.format(picPath,row[1]))
        elif variableFile.CODE_TYPE == "QR":
            qr = qrcode.QRCode(version = None, error_correction = qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
            qr.add_data(tempStr)
            img = qr.make_image()
            nameImg.append('{}QR{}'.format(picPath,index))
            img.save('{}QR{}'.format(picPath,index))

    return nameImg

def xlParmsLIst(xlParms,lbs):
    ''' Creates the key values list to make the dictinoary corresponding
    to the tags in the docx template'''

    for label in range(lbs):
        for param in xlParms:
            variableFile.DICT_KEYS.append('{}{}'.format(param.replace(' ','_'),label+1))

if __name__ == "__main__":
    
    ## TODO This parameters could be given by the user through the interface in a future
    ## TODO: include selection multiple fields to create label 
    xlParms = ["Equipment Model","Serial No", "Work End Date"]
    xlFile = "C:/Users/Javier/Documents/Projects/Docx Labels/Originals/" + "tmpBD1B.xls"
    lblDoc = "C:/Users/Javier/Documents/Projects/Docx Labels/Originals/" + "LabelsJavi2.docx"
    picPath = "C:/Users/Javier/Documents/Projects/Docx Labels/QRpng/"
    tempPath = "C:/Users/Javier/Documents/Projects/Docx Labels/finalTemplates/"
    lbs = variableFile.NUM_LABELS


    xlParmsLIst(xlParms,lbs)
    xlData = readFile(xlFile, xlParms)
    print(xlData)
    # for rows in range(0,len(xlData),lbs):
    #     nameImg = createBarcode(xlData[rows:rows+lbs],picPath)
    #     labelsWord(lblDoc,xlData[rows:rows+lbs],nameImg,rows,tempPath)
