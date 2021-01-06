import os
import threading
import pythoncom
import win32com.client as win32

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk
import pandas as pd

import barcd
import readqr
import variableFile

'''Labels application interface. Main tree:
tk
    MainFrame
        Menu
        NotebookFrame
            LabelFrame
                ExcelFrame
                    FileFrame
                    FilterFrame
                DocxFrame
                    FileFrame
                    FolderFrame
                GenerateFrame
            ScanFrame
            CheckFrame
        BannerFrame
'''

class MainFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.notebook = MainNotebook(self)
        self.banner = BannerFrame(self)
        self.pack(fill = 'both', expand = True)


class MainNotebook(ttk.Notebook):
    def __init__(self,myParent):
        ttk.Notebook.__init__(self,myParent)
        self.pack(fill ='both', expand = True)

        self.labelFrame = LabelFrame(self)
        self.scanFrame = ScanFrame(self)
        self.checkFrame = CheckFrame(self)

        self.add(self.labelFrame, text='Generator')
        self.add(self.scanFrame, text='Scanner')
        self.add(self.checkFrame, text='Checker')

# Classes related with Label Frame

class LabelFrame(tk.Frame):
    '''
    Frame that prensents and launche the main interface for Label creation
    It is divided in multiple frames to allow proper allocation within the Window frame
    '''
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.instFrame = tk.Frame(self)
        self.instLbl = tk.Label(self.instFrame,
            text = 'Instructions',
            justify = tk.LEFT,
            font='5')
        self.filtInstLbl = tk.Label(self.instFrame,
            justify = tk.LEFT,
            text = '\t1.- Select an excel file with the devices information\n\
                2.- Select a word document as a Template for the Labels.\n\
                    2.1- The template parameters should be formated like: {{ Name_Filed }}\n\
                3.- Activate/Deactivate the filter clicking on checkbox "Filter".\n\
                4.- Check Template parameters to filter only by parameters present on the templates\n\
                5.- Select ONE value from first table to filter specific excel column\n\
                6.- Select as many values as needed from second table\n')
        
        self.instFrame.pack(side = 'top',
            fill = 'both',
            padx = (10,10),
            pady = (10,10))
        self.instLbl.pack(side = 'top',
            fill='both',
            anchor = 'w')
        self.filtInstLbl.pack(fill='both', side = 'bottom')

        self.xlFrame = ExcelFrame(self)
        self.docxFrame = DocxFrame(self, self.xlFrame)
        self.genFrame = GenerateFrame(self)

        self.pack(fill = 'both')

    def giveDocxClass(self):
        return self.docxFrame.classFile      


class ExcelFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.granpa = myParent
        self.classFile = None # Reference barcd.py class created
        self.xlFile = FileFrame(
            self,0,'Select excel file',
            (('Excel file','*.xlsx'),('Excel file', '*.xls'),
            ('All files','*.*'),))
        #self.filtFrame = FilterFrame(self)

        self.pack(fill = 'x',
            padx = (10,10),
            pady = (10,10))
    
    def storeClassFile(self,classFile):
        self.classFile = classFile


class DocxFrame(tk.Frame):
    def __init__(self,myParent,xlClass):
        tk.Frame.__init__(self,myParent)
        self.granpa = myParent
        self.classFile = None # Reference barcd.py class created
        self.docxFile = FileFrame(
            self,1,'Select Docx file',
            (('word files','*.docx'),('All files','*.*'),),
            xlClass)
        self.pack(fill='x',
            padx = (10,10),
            pady = (10,10))
        self.filtFrame = FilterFrame(self)

    def storeClassFile(self,classFile):
        self.classFile = classFile


class GenerateFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.topFrame = tk.Frame(self)
        self.botFrame = tk.Frame(self)
        self.gnrBtt = tk.Button(
            self.topFrame, text = 'Generate labels',
            command = lambda: self.generateLbs(),
            state = tk.DISABLED)
        self.gnrLbl = tk.Label(
            self.topFrame,
            text = 'Click on the button to generate the labels')
        self.locFolder = tk.Label(
            self.botFrame,
            text = 'Labels stored in QR_Template folder in Desktop',
            state = tk.DISABLED)

        self.pack(fill='x')
        self.topFrame.pack(side = 'top',
            fill = 'x')
        self.botFrame.pack(side = 'bottom',
            fill = 'x')
        self.gnrLbl.pack(side = 'left',
            anchor = 'w',
            padx = (10,10),
            pady = (10,10))
        self.gnrBtt.pack(side = 'right',
            anchor = 'w',
            padx = (10,10),
            pady = (10,10))
    
    def generateLbs(self):
        docxClass = self.getDocxClass()
        docxClass.labelGenLauncher()
        self.locFolder.pack(side = 'left',
            padx = (10,10),
            pady = (10,10))
        self.gnrBtt['bg'] = 'pale green'
    
    def getDocxClass(self):
        #Redundant
        return self.myParent.giveDocxClass()
    
    def createPB(self):
        # TODO: fix progressBar. "function object has no attribute 'xlDataCaller"
        self.totTmp = len(self.myParent.giveDocxClass.xlDataCaller())/len(self.myParent.giveDocxClass.paramTmp)
        self.prgBar = ttk.Progressbar(
            self.botFrame,
            orient = tk.HORIZONTAL,
            length = 400,
            mode = 'determinate',
            value = 0,
            maximum = self.totTmp)
        self.myParent.giveDocxClass.savePB(self)
        self.prgBar.pack(side = 'left')
    
        
class FileFrame(tk.Frame):
    def __init__(self,myParent,classType,lblText,fileTypes,xlClass = None):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.classFile = None
        ## TODO: the entry is just representing the file, If text displayed by the user is corrected, it should be taken as filepath
        self.filePath = ''
        self.topframe = tk.Frame(self)
        self.midFrame = tk.Frame(self)
        self.btmFrame = tk.Frame(self)
        self.tmpBtt = tk.Button(
            self.midFrame,
            text = 'Browse',
            state = tk.DISABLED,
            command = lambda: self.fileBtw(classType,fileTypes,xlClass))
        if classType == 0:
            self.tmpBtt['state'] = tk.ACTIVE
        self.file = ttk.Entry(
            self.midFrame,
            textvariable = self.filePath,
            width = 80)
        self.frameLbl = ttk.Label(
            self.topframe,
            text=lblText)
        self.errLbl = tk.Label(
            self.btmFrame,
            text ='',
            justify = tk.LEFT)

        self.topframe.pack(fill = 'x')
        self.midFrame.pack(fill = 'x')
        self.btmFrame.pack(fill = 'x', expand = True)

        self.pack(fill = 'x')
        self.frameLbl.pack(
            side = 'top',
            fill = 'both',
            anchor ='nw')
        self.file.pack(side = 'left', anchor = 'nw')
        self.tmpBtt.pack(side = 'top', anchor = 'ne')
        self.errLbl.pack(side = 'bottom', fill = 'both')

    def fileBtw(self,classType,fileTypes,xlClass):
        # excel file browser
        self.filePath = filedialog.askopenfilename(title = "Select file",
            filetypes = fileTypes)
        self.file.delete(0,last = tk.END)
        self.file.insert(0,self.filePath)

        try:
            self.errLbl['text'] = ''
            if classType == 0:
                self.classFile = barcd.XlFile(self.filePath)
                self.myParent.storeClassFile(self.classFile)
                self.myParent.granpa.docxFrame.filtFrame.filterButton['state'] = tk.ACTIVE # Activates filter
                self.myParent.granpa.docxFrame.docxFile.tmpBtt['state'] = tk.ACTIVE
                if self.myParent.granpa.docxFrame.classFile: # Restart xlClass on DocxFrame
                    self.myParent.granpa.docxFrame.classFile = barcd.DocxFile(
                        self.myParent.granpa.docxFrame.docxFile.filePath,
                        self.myParent.classFile)
                self.myParent.granpa.genFrame.gnrBtt['background'] = 'SystemButtonFace'
            elif classType == 1:
                self.classFile = barcd.DocxFile(self.filePath,xlClass.classFile)
                self.myParent.storeClassFile(self.classFile)
                #self.myParent.granpa.genFrame.createPB()
                self.myParent.filtFrame.smpFiltBtt['state'] = tk.ACTIVE
                self.myParent.granpa.genFrame.gnrBtt['state'] = tk.ACTIVE
                self.myParent.granpa.genFrame.gnrBtt['background'] = 'SystemButtonFace'

        except barcd.WrongXlFile:
            self.filePath = ''
            self.errLbl['text'] = 'Wrong file format. Please use file format: xlsx. Or: xls, xlsm, xlsb, odf, ods or odt'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.granpa.docxFrame.filtFrame.filterButton['state'] = tk.DISABLED
        except barcd.WrongDocxFile:
            self.filePath = ''
            self.errLbl['text'] = 'Wrong file format. Please use file format: docx'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
        except barcd.MissingXlFile:
            self.filePath = ''
            self.errLbl['text'] = 'Please insert excel file first'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.granpa.docxFrame.filtFrame.smpFiltBtt['state'] = tk.DISABLED   
        except barcd.EmptyTemplate:
            self.filePath = ''
            self.errLbl['text'] = 'Empty template. Please select a template with at least one field.\
                \nFields should be formated like: {{ Name_Filed }}\
                \nFields must be named as the column on the excel sheet.'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.granpa.docxFrame.filtFrame.smpFiltBtt['state'] = tk.DISABLED
        except barcd.EmbeddedFileError:
            self.filePath = ''
            self.errLbl['text'] = 'Something went wrong. Please check that all included media (e.g. pictures) are embedded in the file.\
            \nPicures should be named as "DummyX" where X corresponds to the number on of the cell'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.granpa.docxFrame.filtFrame.smpFiltBtt['state'] = tk.DISABLED


class FolderFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        #self.pack(fill='x', expand=True)


class FilterFrame(tk.Frame):
    '''Frame with filter options
    Implemented in DocxFrame
    '''
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.filtVar = True
        self.stateSmpFilt = tk.BooleanVar()
        self.interFrame = tk.Frame(self)
        self.bttFrame = tk.Frame(self.interFrame)

        self.filterButton = tk.Checkbutton(
            self.bttFrame, text = 'Filter\t\t\t',
            variable = self.filtVar,
            onvalue = True,
            offvalue = False,
            command = lambda: self.showFilter(self.filtVar),
            state = tk.DISABLED)
        self.smpFiltBtt = tk.Checkbutton(
            self.bttFrame, text = 'Template \n parameters',
            variable = self.stateSmpFilt,
            onvalue = True,
            offvalue = False,
            command = lambda: self.simplyFilter(),
            state = tk.DISABLED)
 
        self.filterFrame = tk.Frame(self.interFrame)
        self.barListParms = tk.Scrollbar(self.filterFrame)
        self.barListValues = tk.Scrollbar(self.filterFrame)
        self.listParms = tk.Listbox(
            self.filterFrame,
            selectmode = 'single',
            yscrollcommand = self.barListParms.set,
            exportselection = 0)
        self.listValues = tk.Listbox(
            self.filterFrame,
            selectmode = 'multiple',
            yscrollcommand = self.barListValues.set,
            exportselection = 0)
        self.barListValues.config(command = self.listValues.yview)
        self.barListParms.config(command = self.listParms.yview)

        self.pack(fill = 'x', side = 'left')
        self.interFrame.pack(side='bottom', fill = 'both')
        self.bttFrame.pack(side = 'left')
        self.filterButton.pack(anchor = 'w', fill = 'x')
        self.smpFiltBtt.pack(side = 'bottom',
            anchor = 'sw',
            after = self.filterButton)

        # Binding
        self.listParms.bind('<<ListboxSelect>>', self.choosenColumn)
        self.listValues.bind('<<ListboxSelect>>', self.choosenValue)

    def showFilter(self, filtVar):
        if filtVar:
            self.myParent.granpa.xlFrame.classFile.filt = True # Activates the filter on barcd.py
            self.filterOptions(filtVar)
            self.filtVar = False
        else:
            self.myParent.granpa.xlFrame.classFile.filt = False
            self.filterOptions(filtVar)
            self.filtVar = True
            self.listValues.delete(0,tk.END)

    def populateLists(self,smply = False):
        '''Creates lists of values to filter. Firts the possible excel columns
        then the values of the column selected.
        '''
        self.listParms.delete(0,tk.END)
        if not smply:
            self.params = self.myParent.granpa.xlFrame.classFile.returnColumns()
        if smply:
            self.params = self.myParent.classFile.paramTmp # Template column parameters
        for param in self.params:
            self.listParms.insert(tk.END, param)              

        self.barListValues.pack(side = 'right', fill = 'x')
        self.listValues.pack(side = 'right')
        self.barListParms.pack(side = 'right', fill = 'x')
        self.listParms.pack(side = 'right')  

    def choosenColumn(self,event):
        '''This function returns the value selected by the user on listParms
        and uses it to present the values on listValues
        '''
        self.listValues.delete(0,tk.END)
        self.filtCol = self.listParms.curselection()[0] # Index column read from template
        self.values = self.myParent.granpa.xlFrame.classFile.returnValues(self.params[self.filtCol])
        self.myParent.granpa.genFrame.gnrBtt['background'] = 'SystemButtonFace'
        for value in self.values:
            self.listValues.insert(tk.END, value)
        
    def choosenValue(self,event):
        temp = list(self.listValues.curselection()) # Returns indexes
        self.filtVal = [self.values[val] for val in temp] # Selected values from filt column
        self.myParent.granpa.xlFrame.classFile.setFilter(self.params[self.filtCol],self.filtVal)
        self.myParent.granpa.genFrame.gnrBtt['background'] = 'SystemButtonFace'
    
    def simplyFilter(self):
        if self.stateSmpFilt.get():
            self.populateLists(True)
        elif not self.stateSmpFilt.get():
            self.populateLists()
        
    def filterOptions(self,filtVar):
        if filtVar:
            self.filterFrame.pack(side = 'right')
            self.populateLists()
        else:
            self.listParms.delete(0,tk.END)
            if self.stateSmpFilt.get():
                self.smpFiltBtt.toggle()
            self.filterFrame.pack_forget()
            self.values = None


class BannerFrame(tk.Frame):
    '''
    Banner that includes:
    MTS logo
    Software version
    Software author
    '''
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill='x', anchor = 'sw', )
    
        imgPath = os.path.join(os.path.dirname(__file__),'mtsHealth.jpg')
        self.versionLbl = tk.Label(self,text = 'Version: 0.1')
        self.authorLbl = tk.Label(self,text = 'By J.Arranz')

        ## NOTE: Consider simply having an image with the specific size when everything is ready
        self.img = Image.open(imgPath)
        self.img = self.img.resize((50,50),Image.ANTIALIAS)
        self.mtsImage = ImageTk.PhotoImage(self.img)
        self.imgLbl = tk.Canvas(self,
            height = 50,
            width = 50)
        self.imgLbl.create_image(25,
            25,
            image = self.mtsImage)
        self.imgLbl.pack(side= 'left')
        self.authorLbl.pack(anchor = 'se', side= 'right')
        self.versionLbl.pack(side = 'bottom', anchor='s')
      
#Classes related with ScanFrame

class ScanFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill = 'both', expand = True)

        self.instFrame = IntrusctLblFrame(self)
        self.reqPumpFrame = ReqPumpFrame(self)
        self.analyticFrame = AnalyticsFrame(self)

        self.instFrame.pack(side = tk.TOP, fill = tk.BOTH, expand = True)
        self.reqPumpFrame.pack(fill = tk.BOTH, expand = True)
        self.analyticFrame.pack(side = tk.BOTTOM, fill = tk.BOTH, expand = True)


class IntrusctLblFrame(tk.Frame):
    '''Frame with instructions and Ready label
    Ready lable changes when the excel file is detected
    '''
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.readyVar = tk.StringVar()
        self.readyVar.set('Open excel')
        self.xl = None

        self.instFrame = tk.Frame(self)
        self.readyFrame = tk.Frame(self)
        self.instLbl = tk.Label(
            self.instFrame,
            text = 'Intructuions')
        self.instTxt = tk.Label(
            self.instFrame, 
            text = 'To be completed')
        self.readLbl = tk.Label(
            self.readyFrame,
            textvariable = self.readyVar,
            font = 25,
            foreground = 'gray')

        self.instFrame.pack(side = tk.LEFT)
        self.readyFrame.pack(side = tk.RIGHT)
        self.instLbl.pack(side = tk.TOP)
        self.instTxt.pack(side = tk.BOTTOM)
        self.readLbl.pack(fill = tk.BOTH)
        self.checkExistXl()
    
    def checkExistXl(self):
        '''Function that constantly checks if excel is open
        If it is, values can be processed, label set to Ready
        If not, display msg when input detected, label set to Red
        '''
        # global changedValue
        try:
            win32.GetActiveObject('Excel.Application')
            if not self.xl:
                self.xl = win32.GetActiveObject('Excel.Application')
                self.processClass = readqr.XlReadWrite(self, self.xl)
                variableFile.changedValue.trace('w', self.processClass.processChanges) # If value changes process function is called
            else:
                self.processClass.checkWb()
        except:
            self.readyVar.set('Open excel')
            self.readLbl.config(foreground = 'gray')

        self.after(1000, self.checkExistXl)
        

class ReqPumpFrame(tk.Frame):
    '''Frame with entry asking for the excel file
    This file should contain the pumps requested by client
    '''
    def __init__(self, myParent):
        tk.Frame.__init__(self,myParent)
        self.filePathEntry = tk.StringVar()
        self.filePathEntry.set('')

        self.lblFrame = tk.Frame(self)
        self.askFileFrame = tk.Frame(self)
        self.lblXlFile = tk.Label(
            self.lblFrame,
            text = 'Requested pumps excel')
        self.fileEntry = tk.Entry(
            self.askFileFrame,
            text = self.filePathEntry,
            width = 80)
        self.browBtt = tk.Button(
            self.askFileFrame,
            text = 'Browse',
            command = lambda: self.fileBtw)

        self.lblFrame.pack(side = tk.TOP, anchor = tk.W)
        self.askFileFrame.pack(anchor = tk.W)
        self.lblXlFile.pack(fill = 'x')
        self.fileEntry.pack(side = tk.LEFT)
        self.browBtt.pack(side = tk.RIGHT)

    def fileBtw(self):
        # excel file browser
        self.filePathEntry.set(
            filedialog.askopenfilename(
                title = "Select file",
                filetypes = (('Excel file', '*.xlsx'), ('Excel file', '*.xls'), ('All files', '*.*'),)))
        self.fileEntry.delete(0,last = tk.END)
        self.fileEntry.insert(0,self.filePathEntry)


class AnalyticsFrame(tk.Frame):
    '''Frame containing the table with pumps
    requested and ready
    '''
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.analLblFrame = tk.Frame(self)
        self.analTblFrame = tk.Frame(self)
        self.analLbl = tk.Label(self.analLblFrame, text = 'Analytics')
        # TODO: Implement table with analytics of current pumps
        self.analTbl = tk.Label(self.analTblFrame, text = 'Here goes the table')

        self.analLblFrame.pack(side = tk.TOP)
        self.analTblFrame.pack(side = tk.BOTTOM)
        self.analLbl.pack(fill = 'x')
        self.analTbl.pack(fill ='x')


class CheckFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill='both', expand=True)
        self.tempLbl = tk.Label(self,
            text = 'Site under construction',
            font = 10,
            highlightcolor='slate gray')
        self.tempLbl.pack(fill='both',
            padx = (10,10),
            pady = (10,10))


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('600x685')
    root.resizable(0,0)
    root.title('MTS Label Manager')
    # TODO: put this on a folder. Media folder or similar with all pictures
    root.iconbitmap(os.path.join(os.path.dirname(__file__),'Icon.ico'))
    # Global variable. Input from user (scanner/keyboard) 
    # Now they come from variableFile.py
    # changedValue = tk.StringVar()
    # changedValue.set('')
    # addressChanged = None
    variableFile.init()
    mainFrame = MainFrame(root)
    root.mainloop()