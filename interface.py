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

class AuxLblFrame(tk.Frame):
    ''' Class especific for the frame that contains the auxiliary label generation
    '''
    ## REVIEW: consider this as general class for all the frames within the main frame.
    ## All share same functions, just needed some special parameters. Maybe usefull if more than 4?
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill='both', expand=True)

        #Ask for a template
        self.auxDocxFilePath = ''
        self.tmpBtt = tk.Button(self, text = 'Browse', command = lambda: self.launchBrwTmp())
        self.auxDocxFile = ttk.Entry(self,textvariable = self.auxDocxFilePath,width = 60)
        self.auxDocxLbl = ttk.Label(self,text='Select auxiliary template Docx file')
        self.errLbldocx = tk.Label(self,text ='',height = 1)

        self.auxDocxLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.auxDocxFile.pack(side = 'left', anchor = 'nw')
        self.tmpBtt.pack(side = 'right', anchor = 'ne')
        self.errLbldocx.pack(side = 'bottom', fill = 'both', after = self.auxDocxLbl)
    
    def launchBrwTmp(self):
        # word file browser
        self.auxDocxFilePath = filedialog.askopenfilename(title = "Select auxiliary template file", filetypes = (('word files','*.docx'),('All files','*.*'),))
        self.auxDocxFile.select_clear()
        self.auxDocxFile.insert(0,self.auxDocxFilePath)



class LabelFrame(tk.Frame):
    '''
    Frame that prensents and launche the main interface for Label creation
    It is divided in multiple frames to allow proper allocation within the Window frame
    '''
    def __init__(self,myParent):
        tk.Frame.__init__(self,myParent)
        self.pack(fill='both', expand=True)

        ## REVIEW: this should be selected by user in the future.
        self.lbs = variableFile.NUM_LABELS
        self.xlParms = variableFile.XLS_PARAM
        barcd.xlParmsLIst(self.xlParms,self.lbs)


        # Organization frames
        self.topFrame = tk.Frame(self)
        self.midFrame = tk.Frame(self)
        self.auxFrame = tk.Frame(self)
        self.btmFrame = tk.Frame(self, height = 50)
        self.btmAuxFrame = tk.Frame(self,heigh = 50)

        self.topFrame.pack(fill = 'x',anchor = 'nw')
        self.midFrame.pack(fill = 'x',after = self.topFrame)
        self.btmFrame.pack(fill = 'x', after = self.midFrame)
        self.auxFrame.pack(fill = 'x', after = self.btmFrame)
        self.btmAuxFrame.pack(fill = 'x',after = self.auxFrame)

        #Ask for a main template
        self.docxFilePath = ''
        self.tmpBtt = tk.Button(self.topFrame, text = 'Browse', command = lambda: self.launchBrwTmp())
        self.docxFile = ttk.Entry(self.topFrame,textvariable = self.docxFilePath,width = 60)
        self.docxLbl = ttk.Label(self.topFrame,text='Select main template Docx file')
        self.errLbldocx = tk.Label(self.topFrame,text ='',height = 1)

        self.docxLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.docxFile.pack(side = 'left', anchor = 'nw')
        self.tmpBtt.pack(side = 'right', anchor = 'ne')
        self.errLbldocx.pack(side = 'bottom', fill = 'both', after = self.docxLbl)

        # Ask for file name
        self.exlFilePath = ''
        self.brwsBtt = tk.Button(self.midFrame, text = 'Browse', command = lambda : self.launchBrw())
        self.exlFile = ttk.Entry(self.midFrame,textvariable = self.exlFilePath,width = 60)
        self.pthLbl = ttk.Label(self.midFrame,text='Select excel file')
        self.errLblxl = tk.Label(self.midFrame,text='',height = 1)

        self.pthLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.exlFile.pack(side = 'left', anchor = 'nw')
        self.brwsBtt.pack(side = 'right', anchor = 'ne')
        self.errLblxl.pack(side = 'bottom',fill = 'both', after = self.pthLbl)

        # Generate templates
        self.gnrBtt = tk.Button(self.btmFrame, text = 'Generate labels', command = lambda: self.generateLbs(), bg = 'orange')
        self.gnrLbl = tk.Label(self.btmFrame, text = 'Click on the button to generate the labels')

        self.gnrLbl.pack(side = 'left', anchor = 'w')
        self.gnrBtt.pack(side = 'right', anchor = 'w')

        #Ask for a Auxiliary template
        self.docxAuxFilePath = ''
        self.tmpAuxBtt = tk.Button(self.auxFrame, text = 'Browse', command = lambda: self.launchAuxBrwTmp())
        self.docxAuxFile = ttk.Entry(self.auxFrame,textvariable = self.docxAuxFilePath,width = 60)
        self.docxAuxLbl = ttk.Label(self.auxFrame,text='Select auxiliary template Docx file')
        self.errAuxLbldocx = tk.Label(self.auxFrame,text ='',height = 1)

        self.docxAuxLbl.pack(side = 'top', fill = 'both', anchor ='nw')
        self.docxAuxFile.pack(side = 'left', anchor = 'nw')
        self.tmpAuxBtt.pack(side = 'right', anchor = 'ne')
        self.errAuxLbldocx.pack(side = 'bottom', fill = 'both', after = self.docxAuxLbl)

        # Generate auxiliary templates
        self.gnrAuxBtt = tk.Button(self.btmAuxFrame, text = 'Generate auxiliary labels', command = lambda: self.generateAuxLbs(), bg = 'orange')
        self.gnrAuxLbl = tk.Label(self.btmAuxFrame, text = 'Click on the button to generate auxiliary labels')

        self.gnrAuxLbl.pack(side = 'left', anchor = 'w')
        self.gnrAuxBtt.pack(side = 'right', anchor = 'w')



    def launchBrw(self):
        # excel file browser
        self.exlFilePath = filedialog.askopenfilename(title = "Select file", filetypes = (('Excel file','*.xlsx'),('Excel file', '*.xls'),('All files','*.*'),))
        self.exlFile.delete(0,last = tk.END)
        self.exlFile.insert(0,self.exlFilePath)
        try:
            self.errLblxl['text'] = ''
            self.xlData = barcd.readFile(self.exlFilePath, self.xlParms)
        except barcd.WrongXlFile:
            self.exlFilePath = ''
            self.errLblxl['text'] = 'Wrong file format. Please use file format: xls, xlsx, xlsm, xlsb, odf, ods or odt'
            self.exlFile.delete(0,last = tk.END)
            self.exlFile.insert(0,self.exlFilePath)

    
    def launchBrwTmp(self):
        # word file browser
        self.docxFilePath = filedialog.askopenfilename(title = "Select main template file", filetypes = (('word files','*.docx'),('All files','*.*'),))
        self.docxFile.delete(0,last = tk.END)
        self.docxFile.insert(0,self.docxFilePath)
        ## TODO: implement method to read file after selecting it to create menu to check selected parameters
    
    def launchAuxBrwTmp(self):
        # word file browser
        self.auxDocxFilePath = filedialog.askopenfilename(title = "Select auxiliary template file", filetypes = (('word files','*.docx'),('All files','*.*'),))
        self.docxAuxFile.delete(0,last = tk.END)
        self.docxAuxFile.insert(0,self.auxDocxFilePath)
    
    def generateLbs(self):
        # Calls main functions to generate barcd

        ## TODO: this folders should be created (and destroyed) by the programm automaticaly.
        picPath = "C:/Users/Javier/Documents/Projects/Docx Labels/QRpng/"
        tempPath = "C:/Users/Javier/Documents/Projects/Docx Labels/finalTemplates/"

        # Main functions call with file checking
        # REVIEW: review this whole for loop. Probably it would be better implemented in a single function
        try:
            self.errLbldocx['text'] = ''
            for rows in range(0,len(self.xlData),self.lbs):
                nameImg = barcd.createBarcode(self.xlData[rows:rows+self.lbs],picPath)
                barcd.labelsWord(self.docxFilePath,self.xlData[rows:rows+self.lbs],nameImg,rows,tempPath)
            self.gnrBtt['bg'] = 'green'
            self.gnrLbl['text'] = 'Labels generated!'

        except barcd.WrongDocxFile:
             self.errLbldocx['text']= 'Wrong file format. Please use file format: docx'

    def generateAuxLbs(self):
        ## TODO: this folders should be created (and destroyed) by the programm automaticaly.
        picPath = "C:/Users/Javier/Documents/Projects/Docx Labels/QRpng/"
        tempPath = "C:/Users/Javier/Documents/Projects/Docx Labels/finalTemplates/"

        # Main functions call with file checking
        # REVIEW: review this whole for loop. Probably it would be better implemented in a single function

        ## TODO: paramFilt and model should be selected by user
        ## REVIEW: this function has not been tested yet
        try:
            self.errAuxLbldocx['text'] = ''
            barcd.auxLabelWord(self.docxAuxFilePath,self.xlData,picPath,tempPath,paramFilt = 'Equipment Model', model = 'BODYGUARD 323')
            self.gnrAuxBtt['bg'] = 'green'
            self.gnrAuxLbl['text'] = 'Labels generated!'
        except barcd.WrongDocxFile:
             self.errAuxLbldocx['text']= 'Wrong file format. Please use file format: docx'

        



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

        self.imgLbl = tk.Canvas(self, height = 50, width = 50)
        
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