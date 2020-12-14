import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk
import variableFile
import os
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
        self.classFile = None
        self.filtVar = tk.BooleanVar()
        self.xlFile = FileFrame(self,0,'Select excel file',(('Excel file','*.xlsx'),('Excel file', '*.xls'),('All files','*.*'),))
        self.filtFrame = FilterFrame(self)
        self.filterButton = tk.Checkbutton(self, text = 'Filter', variable = self.filtVar, onvalue = True, offvalue = False, command = lambda: self.filtFrame.showFilter(self.filtVar,self))

        self.pack(fill = 'x')
        #self.filterButton.pack()
    
    
    def storeClassFile(self,classFile):
        self.classFile = classFile


class DocxFrame(tk.Frame):
    def __init__(self,myParent,xlClass):
        tk.Frame.__init__(self,myParent)
        self.classFile = None
        self.docxFile = FileFrame(self,1,'Select Docx file',(('word files','*.docx'),('All files','*.*'),),xlClass)
        #self.classFile = self.docxFile.returnClass()
        self.pack(fill='x')

    def storeClassFile(self,classFile):
        self.classFile = classFile


class GenerateFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.myParent = myParent
        self.gnrBtt = tk.Button(self, text = 'Generate labels', command = lambda: self.generateLbs(), bg = 'orange')
        self.gnrLbl = tk.Label(self, text = 'Click on the button to generate the labels')

        self.pack(fill='x')
        self.gnrLbl.pack(side = 'left', anchor = 'w')
        self.gnrBtt.pack(side = 'right', anchor = 'w')
    
    def generateLbs(self):
        docxClass = self.getDocxClass()
        docxClass.labelGenLauncher()
    
    def getDocxClass(self):
        return self.myParent.giveDocxClass()

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
            elif classType == 1:
                self.classFile = barcd.DocxFile(self.filePath,xlClass.classFile)
                self.myParent.storeClassFile(self.classFile)

        except barcd.WrongXlFile:
            self.filePath = ''
            self.errLbl['text'] = 'Wrong file format. Please use file format: xls, xlsx, xlsm, xlsb, odf, ods or odt'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
        except barcd.WrongDocxFile:
            self.filePath = ''
            self.errLbl['text'] = 'Wrong file format. Please use file format: docx'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)
        except barcd.MissingXlFile:
            self.errLbl['text'] = 'Please insert excel file first'
            self.file.delete(0,last = tk.END)
            self.file.insert(0,self.filePath)   


class FolderFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        #self.pack(fill='x', expand=True)

class FilterFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.lo = tk.Label(self, text = 'Click on the button to generate the labels')
    
    def showFilter(self, filtVar,parent):
        if filtVar:
            self.pack()
            self.lo.pack(fill = 'x')
        else:
            self.lo.destroy()
            parent.pack()


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