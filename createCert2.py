import openpyxl
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfFileMerger
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import ttk
from tkinter import *

class CPA_Cert_Generator (tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry('580x470')
        self.configure(background='#FAEBD7')
        self.title('CPA Certificate Generator')

        self.buttonFont = ('courier', 12, 'normal')
        self.labelFont = ('courier', 12, 'bold')
        self.pathFont = ('courier', 12, 'normal')
        self.logFont = ('courier', 10, 'normal')

        self.logBox = scrolledtext.ScrolledText(self, wrap = tk.WORD, width = 60, height = 10, font = self.logFont)
        self.logBox.place(x=40, y=278)
        self.logBox.insert(tk.INSERT, 'Please select files \n')
        self.logBox.configure(state='disabled')

        self.selectedTemplate = Label(self, text='', bg='#FAEBD7', font=self.pathFont)
        self.selectedTemplate.place(x=230, y=78)

        self.selectedLifterData = Label(self, text='', bg='#FAEBD7', font=self.pathFont)
        self.selectedLifterData.place(x=260, y=158)

        self.templateButton = Button(self, text='Select Template File', bg='#FAEBD7', font=self.buttonFont, command=self.selectTemplate)
        self.templateButton.place(x=40, y=38)

        Label(self, text='Selected Template:', bg='#FAEBD7', font=self.labelFont).place(x=40, y=78)

        self.lifterDataButton = Button(self, text='Select Lifter Data', bg='#FAEBD7', font=self.buttonFont, command=self.selectLifterData)
        self.lifterDataButton.place(x=40, y=118)

        Label(self, text='Selected Lifter Data:', bg='#FAEBD7', font=self.labelFont).place(x=40, y=158)

        self.runButton = Button(self, text='Generate Certificate', bg='#FAEBD7', font=self.buttonFont, command=self.process_data)
        self.runButton.place(x=40, y=198)

    def selectLifterData(self):
        inputdata = filedialog.askopenfilename(filetypes=[('Lifter data', '*.xlsx')])
        self.selectedLifterData.configure(text = inputdata)
        self.logBox.configure(state='normal')
        self.logBox.insert(tk.INSERT, 'selected input data:\n{0}\n'.format(inputdata))
        self.logBox.configure(state='disabled')
        self.logBox.see("end")

    def selectTemplate(self):
        template = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
        self.selectedTemplate.configure(text = template)
        self.logBox.configure(state='normal')
        self.logBox.insert(tk.INSERT, 'selected template:\n{0}\n'.format(template))
        self.logBox.configure(state='disabled')
        self.logBox.see("end")

    def process_data(self):
        print("process")


if __name__ == '__main__':
    app = CPA_Cert_Generator()
    app.mainloop()