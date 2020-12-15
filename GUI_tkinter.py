import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk
import variableFile
import os
import barcd
import threading

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
        self.classFile = None
        self.xlFile = FileFrame(self,0,'Select excel file',(('Excel file','*.xlsx'),('Excel file', '*.xls'),('All files','*.*'),))
        self.filtFrame = FilterFrame(self)

        self.pack(fill = 'x')
    
    
    def storeClassFile(self,classFile):
        self.classFile = classFile



class DocxFrame(tk.Frame):
    def __init__(self,myParent,xlClass):
        tk.Frame.__init__(self,myParent)
        self.granpa = myParent
        self.classFile = None
        self.docxFile = FileFrame(self,1,'Select Docx file',(('word files','*.docx'),('All files','*.*'),),xlClass)
        self.pack(fill='x')

    def storeClassFile(self,classFile):
        self.classFile = classFile


class GenerateFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.topFrame = tk.Frame(self)
        self.botFrame = tk.Frame(self)
        self.gnrBtt = tk.Button(self.topFrame, text = 'Generate labels', command = lambda: self.generateLbs(), bg = 'orange')
        self.gnrLbl = tk.Label(self.topFrame, text = 'Click on the button to generate the labels')
        self.locFolder = tk.Label(self.botFrame, text = 'Labels stored in QR_Template folder in Desktop', state = tk.DISABLED)


        self.pack(fill='x')
        self.topFrame.pack(side = 'top', fill = 'x')
        self.botFrame.pack(side = 'bottom', fill = 'x')
        self.gnrLbl.pack(side = 'left', anchor = 'w')
        self.gnrBtt.pack(side = 'right', anchor = 'w')
    
    def generateLbs(self):
        docxClass = self.getDocxClass()
        docxClass.labelGenLauncher()
        self.locFolder.pack(side = 'left')
    
    def getDocxClass(self):
        #Redundant
        return self.myParent.giveDocxClass()
    
    def createPB(self):
        # TODO: fix progressBar. "function object has no attribute 'xlDataCaller"
        self.totTmp = len(self.myParent.giveDocxClass.xlDataCaller())/len(self.myParent.giveDocxClass.paramTmp)
        self.prgBar = ttk.Progressbar(self.botFrame,orient = tk.HORIZONTAL, length = 400, mode = 'determinate', value = 0, maximum = self.totTmp)
        self.myParent.giveDocxClass.savePB(self)
        self.prgBar.pack(side = 'left')
    
        

class FileFrame(tk.Frame):
    def __init__(self,myParent,classType,lblText,fileTypes,xlClass = None):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.classFile = None
        ## TODO: the entry is just representing the file, If text displayed by the user is corrected, it should be taken as filepath
        self.filePath = ''
        self.tmpBtt = tk.Button(self, text = 'Browse', command = lambda: self.fileBtw(classType,fileTypes,xlClass))
        self.file = ttk.Entry(self,textvariable = self.filePath,width = 60)
        self.frameLbl = ttk.Label(self,text=lblText)
        self.errLbl = tk.Label(self,text ='',height = 1)

        self.pack(fill = 'x')
        self.frameLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.file.pack(side = 'left', anchor = 'nw')
        self.tmpBtt.pack(side = 'right', anchor = 'ne')
        self.errLbl.pack(side = 'bottom', after = self.file, fill = 'both')

    
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
            elif classType == 1:
                self.classFile = barcd.DocxFile(self.filePath,xlClass.classFile)
                self.myParent.storeClassFile(self.classFile)
                self.myParent.granpa.genFrame.createPB()

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
            self.errLbl['text'] = 'Please insert excel file first'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
            self.myParent.filtFrame.filterButton['state'] = tk.DISABLED   


class FolderFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        #self.pack(fill='x', expand=True)

class FilterFrame(tk.Frame):
    # TODO: filter button should be active only when there is an excel loaded.
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.filtVar = True
        self.smpFltBtt = False
        self.bttFrame = tk.Frame(self)
        self.filterButton = tk.Checkbutton(self.bttFrame, text = 'Filter                 ', variable = self.filtVar, onvalue = True, offvalue = False, command = lambda: self.showFilter(self.filtVar), state = tk.DISABLED)
        self.smpFiltBtt = tk.Checkbutton(self.bttFrame, text = 'Template \n parameters', variable = self.smpFltBtt, onvalue = True, offvalue = False, state = tk.DISABLED)


        self.filterFrame = tk.Frame(self)
        self.barListParms = tk.Scrollbar(self.filterFrame)
        self.barListValues = tk.Scrollbar(self.filterFrame)
        self.listParms = tk.Listbox(self.filterFrame,selectmode = 'browse', yscrollcommand = self.barListParms.set, )
        self.listValues = tk.Listbox(self.filterFrame, selectmode = 'browse', yscrollcommand = self.barListValues.set)
        self.barListValues.config(command = self.listValues.yview)
        self.barListParms.config(command = self.listParms.yview)

        self.pack(fill = 'x',side = 'left')
        self.bttFrame.pack(side = 'left')
        self.filterButton.pack(anchor = 'w', fill = 'x')
        self.smpFiltBtt.pack(side = 'bottom',anchor = 'sw', after = self.filterButton)

        # TODO: binding not working. Fix.
        self.listParms.bind('<<ListboxSelect>>',self.choosenColumn())

    def showFilter(self, filtVar):
        if filtVar:
            self.filterOptions(filtVar)
            self.filtVar = False
        else:
            self.filterOptions(filtVar)
            self.filtVar = True

    def populateLists(self):
        '''Creates lists of values to filter. Firts the possible excel columns
        then the values of the column selected.
        '''
        self.listParms.delete(0,tk.END)
        self.params = self.myParent.classFile.returnColumns()
        for param in self.params:
            self.listParms.insert(tk.END, param)              

        self.barListValues.pack(side = 'right',fill = 'both')
        self.listValues.pack(side = 'right')
        self.barListParms.pack(side = 'right', fill = 'both')
        self.listParms.pack(side = 'right')  

        

    def choosenColumn(self):
        '''This function returns the value selected by the user on listParms
        and uses it to present the values on listValues
        '''
        # TODO: not working. Fix.
        #self.values = self.myParent.classFile.returnValues(self.listParms.curselection())
        #for value in self.values:
        #    self.listValues.insert(tk.END, value)
        pass

        
        
    
    def filterOptions(self,filtVar):
        # REVIEW: this is not correctly implemented
        # NOTE: think using a list instead of buttons, sorter.
        # REVIEW: Not deleting a
        if filtVar:
            self.filterFrame.pack(side = 'right')
            self.populateLists()
        else:
            self.filterFrame.pack_forget()
    
    

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
        

        ## TODO: Find sol for pictures in .exe python files
        imgPath = "C:\\Users\\Javier\\Documents\\Projects\\Docx Labels\\Icons\\mtsHealth.jpg"
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
    root.geometry('500x500')
    root.minsize(500,500)
    root.title('MTS Label Manager')
    mainFrame = MainFrame(root)
    root.mainloop()