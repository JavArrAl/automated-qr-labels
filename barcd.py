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


class WrongXlFile(Exception):
    pass

class WrongDocxFile(Exception):
    pass


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

def auxLabelWord(auxlblDoc,xlData,picPath,tempPath,paramFilt = 'Equipment Model', model = 'BODYGUARD 323'):
    '''Invokes methods to create lables using only the models selected
    '''
    ## NOTE: This has not been tested.
    lbs = variableFile.NUM_LABELS_AUX
    xlData = xlData[xlData[paramFilt] == model]
    for rows in range(0,len(xlData),lbs):
        nameImg = createBarcode(xlData[rows:rows+lbs],picPath)
        labelsWord(auxlblDoc,xlData[rows:rows+lbs],nameImg,rows,tempPath,'Aux_Templates')
 

def labelsWord(lblDoc,xlData,nameImg,numTemp,tempPath,nameTmp = 'Templates'):
    ''' Creates a dictionary matching values from the excel introduced by the user
    to the tags in the docx template.
    '''
    ## TODO: Dynamic generation of table cells based on number of lbsl necessary.

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
