import variableFile
import os
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk

import barcd

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


class LabelFrame(tk.Frame):
    '''
    Frame that prensents and launche the main interface for Label creation
    It is divided in multiple frames to allow proper allocation within the Window frame
    '''
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.xlFrame = ExcelFrame(self)
        self.docxFrame = DocxFrame(self,self.xlFrame)
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
        self.filtFrame = FilterFrame(self)

        self.pack(fill = 'x', padx = (10,10), pady = (10,10))
    
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
        self.pack(fill='x', padx = (10,10), pady = (10,10))

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
            bg = 'orange')
        self.gnrLbl = tk.Label(
            self.topFrame, text = 'Click on the button to generate the labels')
        self.locFolder = tk.Label(
            self.botFrame, text = 'Labels stored in QR_Template folder in Desktop',
            state = tk.DISABLED)

        self.pack(fill='x')
        self.topFrame.pack(side = 'top', fill = 'x')
        self.botFrame.pack(side = 'bottom', fill = 'x')
        self.gnrLbl.pack(side = 'left', anchor = 'w', padx = (10,10), pady = (10,10))
        self.gnrBtt.pack(side = 'right', anchor = 'w', padx = (10,10), pady = (10,10))
    
    def generateLbs(self):
        docxClass = self.getDocxClass()
        docxClass.labelGenLauncher()
        self.locFolder.pack(side = 'left', padx = (10,10), pady = (10,10))
    
    def getDocxClass(self):
        #Redundant
        return self.myParent.giveDocxClass()
    
    def createPB(self):
        # TODO: fix progressBar. "function object has no attribute 'xlDataCaller"
        self.totTmp = len(self.myParent.giveDocxClass.xlDataCaller())/len(self.myParent.giveDocxClass.paramTmp)
        self.prgBar = ttk.Progressbar(
            self.botFrame,orient = tk.HORIZONTAL,
            length = 400, mode = 'determinate',
            value = 0, maximum = self.totTmp)
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
            self.midFrame, text = 'Browse',
            command = lambda: self.fileBtw(classType,fileTypes,xlClass))
        self.file = ttk.Entry(self.midFrame,textvariable = self.filePath,width = 80)
        self.frameLbl = ttk.Label(self.topframe,text=lblText)
        self.errLbl = tk.Label(self.btmFrame,text ='', justify = tk.LEFT)

        self.topframe.pack(fill = 'x')
        self.midFrame.pack(fill = 'x')
        self.btmFrame.pack(fill = 'x', expand = True)

        self.pack(fill = 'x')
        self.frameLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.file.pack(side = 'left', anchor = 'nw')
        self.tmpBtt.pack(side = 'top', anchor = 'ne')
        self.errLbl.pack(side = 'bottom', fill = 'both')

    def fileBtw(self,classType,fileTypes,xlClass):
        # excel file browser
        self.filePath = filedialog.askopenfilename(title = "Select file", filetypes = fileTypes)
        self.file.delete(0,last = tk.END)
        self.file.insert(0,self.filePath)

        try:
            self.errLbl['text'] = ''
            if classType == 0:
                self.classFile = barcd.XlFile(self.filePath)
                self.myParent.storeClassFile(self.classFile)
                self.myParent.filtFrame.filterButton['state'] = tk.ACTIVE # Activates filter
                if self.myParent.granpa.docxFrame.classFile: # Restart xlClass on DocxFrame
                    self.myParent.granpa.docxFrame.classFile = barcd.DocxFile(
                        self.myParent.granpa.docxFrame.docxFile.filePath,
                        self.myParent.classFile)
            elif classType == 1:
                self.classFile = barcd.DocxFile(self.filePath,xlClass.classFile)
                self.myParent.storeClassFile(self.classFile)
                #self.myParent.granpa.genFrame.createPB()
                self.myParent.granpa.xlFrame.filtFrame.smpFiltBtt['state'] = tk.ACTIVE

        except barcd.WrongXlFile:
            self.filePath = ''
            self.errLbl['text'] = 'Wrong file format. Please use file format: xls, xlsx, xlsm, xlsb, odf, ods or odt'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.filtFrame.filterButton['state'] = tk.DISABLED
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
            self.myParent.granpa.xlFrame.filtFrame.smpFiltBtt['state'] = tk.DISABLED   
        except barcd.EmptyTemplate:
            self.filePath = ''
            self.errLbl['text'] = 'Empty template. Please select a template with at least one field.\
                \nFields should be formated like: {{ Name_Filed }}\
                \nFields must be named as the column on the excel sheet.'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.granpa.xlFrame.filtFrame.smpFiltBtt['state'] = tk.DISABLED
        except barcd.EmbeddedFileError:
            self.filePath = ''
            self.errLbl['text'] = 'Something went wrong. Please check that all included media (e.g. pictures) are embedded in the file.\
            \nPicures should be named as "DummyX" where X corresponds to the number on of the cell'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.granpa.xlFrame.filtFrame.smpFiltBtt['state'] = tk.DISABLED


class FolderFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        #self.pack(fill='x', expand=True)


class FilterFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.filtVar = True
        self.stateSmpFilt = tk.BooleanVar()
        self.bttFrame = tk.Frame(self)
        self.filterButton = tk.Checkbutton(
            self.bttFrame, text = 'Filter\t\t\t',
            variable = self.filtVar, onvalue = True,
            offvalue = False, command = lambda: self.showFilter(self.filtVar),
            state = tk.DISABLED)
        self.smpFiltBtt = tk.Checkbutton(
            self.bttFrame, text = 'Template \n parameters',
            variable = self.stateSmpFilt, onvalue = True,
            offvalue = False, command = lambda: self.simplyFilter(),
            state = tk.DISABLED)


        self.filterFrame = tk.Frame(self)
        self.barListParms = tk.Scrollbar(self.filterFrame)
        self.barListValues = tk.Scrollbar(self.filterFrame)
        self.listParms = tk.Listbox(
            self.filterFrame,selectmode = 'single',
            yscrollcommand = self.barListParms.set,
            exportselection = 0)
        self.listValues = tk.Listbox(
            self.filterFrame, selectmode = 'multiple',
            yscrollcommand = self.barListValues.set,
            exportselection = 0)
        self.barListValues.config(command = self.listValues.yview)
        self.barListParms.config(command = self.listParms.yview)
        self.pack(fill = 'x',side = 'left')
        self.bttFrame.pack(side = 'left')
        self.filterButton.pack(anchor = 'w', fill = 'x')
        self.smpFiltBtt.pack(side = 'bottom',anchor = 'sw', after = self.filterButton)

        # Binding
        self.listParms.bind('<<ListboxSelect>>', self.choosenColumn)
        self.listValues.bind('<<ListboxSelect>>', self.choosenValue)

    def showFilter(self, filtVar):
        if filtVar:
            self.myParent.classFile.filt = True # Activates the filter on barcd.py
            self.filterOptions(filtVar)
            self.filtVar = False
        else:
            self.myParent.classFile.filt = False
            self.filterOptions(filtVar)
            self.filtVar = True
            self.listValues.delete(0,tk.END)

    def populateLists(self,smply = False):
        '''Creates lists of values to filter. Firts the possible excel columns
        then the values of the column selected.
        '''
        self.listParms.delete(0,tk.END)
        if not smply:
            self.params = self.myParent.classFile.returnColumns()
        if smply:
            self.params = self.myParent.granpa.docxFrame.classFile.paramTmp # Template column parameters
        for param in self.params:
            self.listParms.insert(tk.END, param)              

        self.barListValues.pack(side = 'right',fill = 'both')
        self.listValues.pack(side = 'right')
        self.barListParms.pack(side = 'right', fill = 'both')
        self.listParms.pack(side = 'right')  

    def choosenColumn(self,event):
        '''This function returns the value selected by the user on listParms
        and uses it to present the values on listValues
        '''
        self.listValues.delete(0,tk.END)
        self.filtCol = self.listParms.curselection()[0] # Index column read from template
        self.values = self.myParent.classFile.returnValues(self.params[self.filtCol])
        for value in self.values:
            self.listValues.insert(tk.END, value)
        
    def choosenValue(self,event):
        temp = list(self.listValues.curselection()) # Returns indexes
        self.filtVal = [self.values[val] for val in temp] # Selected values from filt column
        self.myParent.classFile.setFilter(self.params[self.filtCol],self.filtVal)
    
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
      

class ScanFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        #self.pack(fill='both', expand=True)


class CheckFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        #self.pack(fill='both', expand=True)


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
        self.imgLbl = tk.Canvas(self, height = 50, width = 50)
        self.imgLbl.create_image(25,25,image = self.mtsImage)
        self.imgLbl.pack(side= 'left')
        self.authorLbl.pack(anchor = 'se', side= 'right')
        self.versionLbl.pack(side = 'bottom',anchor='s')


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('600x500')
    root.resizable(0,0)
    root.title('MTS Label Manager')
    root.iconbitmap(os.path.join(os.path.dirname(__file__),'Icon.ico'))
    mainFrame = MainFrame(root)
    root.mainloop()