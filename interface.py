import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk
import variableFile
import os
import barcd


class MainFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.notebook = MainNotebook(self)
        self.banner = BannerFrame(self)

        self.pack(expand = True, fill = 'both')

class MainNotebook(ttk.Notebook):
    def __init__(self,myParent):
        ttk.Notebook.__init__(self,myParent)
        self.pack( expand = True, fill ='both' )

        self.labelFrame = LabelFrame(self)
        self.scanFrame = ScanFrame(self)
        self.checkFrame = CheckFrame(self)

        self.add(self.labelFrame, text='Generator')
        self.add(self.scanFrame, text='Scanner')
        self.add(self.checkFrame, text='Checker')

class LabelFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill='both', expand=True)

        # Organization frames
        self.topFrame = tk.Frame(self)
        self.midFrame = tk.Frame(self)
        self.btmFrame = tk.Frame(self, height = 50)

        self.topFrame.pack(fill = 'x',anchor = 'nw')
        self.midFrame.pack(fill = 'x',after = self.topFrame)
        self.btmFrame.pack(fill = 'x', after = self.midFrame)

        # Ask for file name
        self.exlFilePath = ''
        self.brwsBtt = tk.Button(self.topFrame, text = 'Browse', command = lambda : self.launchBrw())
        self.exlFile = ttk.Entry(self.topFrame,textvariable = self.exlFilePath,width = 60)
        self.pthLbl = ttk.Label(self.topFrame,text='Select excel file')

        self.pthLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.exlFile.pack(side = 'left', anchor = 'nw')
        self.brwsBtt.pack(side = 'right', anchor = 'ne')

        #Ask for a template
        self.docxFilePath = ''
        self.tmpBtt = tk.Button(self.midFrame, text = 'Browse', command = lambda: self.launchBrwTmp())
        self.docxFile = ttk.Entry(self.midFrame,textvariable = self.exlFilePath,width = 60)
        self.docxLbl = ttk.Label(self.midFrame,text='Select Docx file')

        self.docxLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.docxFile.pack(side = 'left', anchor = 'nw')
        self.tmpBtt.pack(side = 'right', anchor = 'ne')

        # Generate templates
        self.gnrBtt = tk.Button(self.btmFrame, text = 'Generate labels', command = lambda: self.generateLbs(), bg = 'red')
        self.gnrLbl = tk.Label(self.btmFrame, text = 'Click on the button to generate the labels')

        self.gnrLbl.pack(side = 'left', anchor = 'w')
        self.gnrBtt.pack(side = 'right')


    def launchBrw(self):
        # excel file browser
        self.exlFilePath = filedialog.askopenfilename(initialdir = os.getcwd() ,title = "Select file", filetypes = (('excel files','*.xlsx'),('All files','*.*'),))
        self.exlFile.select_clear()
        self.exlFile.insert(0,self.exlFilePath)
    
    def launchBrwTmp(self):
        # word file browser
        self.docxFilePath = filedialog.askopenfilename(initialdir = os.getcwd() ,title = "Select file", filetypes = (('word files','*.docx'),('All files','*.*'),))
        self.docxFile.select_clear()
        self.docxFile.insert(0,self.docxFilePath)
    
    def generateLbs(self):
        # Calls main functions to generate barcd
        ## TODO: needs mayor error handling. E.x: file paths empty, not working files, wrong file extensions...
        lbs = variableFile.NUM_LABELS
        xlParms = variableFile.XLS_PARAM
        ## TODO: this folders should be created (and destroyed) by the programm automaticaly.
        picPath = "C:/Users/Javier/Documents/Projects/Docx Labels/QRpng/"
        tempPath = "C:/Users/Javier/Documents/Projects/Docx Labels/finalTemplates/"

        barcd.xlParmsLIst(xlParms,lbs)
        xlData = barcd.readFile(self.exlFilePath, xlParms)

        for rows in range(0,len(xlData),lbs):
            nameImg = barcd.createBarcode(xlData[rows:rows+lbs],picPath)
            barcd.labelsWord(self.docxFilePath,xlData[rows:rows+lbs],nameImg,rows,tempPath)


class ScanFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill='both', expand=True)

class CheckFrame(tk.Frame):
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill='both', expand=True)

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
        self.authorLbl = tk.Label(self,text = 'By Javier Arranz')

        ## NOTE: Consider simply having an image with the specific size when everything is ready
        self.img = Image.open(imgPath)
        self.img = self.img.resize((50,50),Image.ANTIALIAS)
        self.mtsImage = ImageTk.PhotoImage(self.img)

        self.imgLbl = tk.Canvas(self, height = 55, width = 55)
        
        self.imgLbl.create_image(25,25,image = self.mtsImage)
        self.imgLbl.pack(side= 'left')
        self.authorLbl.pack(anchor = 'se', side= 'right')
        self.versionLbl.pack(side = 'bottom',anchor='s')
        

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('500x500')
    root.title('MTS Label Manager')
    mainFrame = MainFrame(root)
    root.mainloop()