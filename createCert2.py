import os
import openpyxl
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import ttk
from tkinter import *

class CPA_Cert_Generator (tk.Tk):
    def __init__(self):
        """Initialize the GUI application with buttons, labels, and a text box for logging."""

        super().__init__()

        # Set the initial values of the template and lifterData attributes to None.
        self.template = None
        self.lifterData = None

        # Set the size, background color, and title of the window.
        self.geometry('580x470')
        self.configure(background='#FAEBD7')
        self.title('CPA Certificate Generator')

        # Set the fonts for buttons, labels, and file path displays.
        self.buttonFont = ('courier', 12, 'normal')
        self.labelFont = ('courier', 12, 'bold')
        self.pathFont = ('courier', 12, 'normal')
        self.logFont = ('courier', 10, 'normal')

        # Set up the logging text box and configure it with tags for different message types.
        self.logBox = scrolledtext.ScrolledText(self, wrap=tk.WORD, width=60, height=10, font=self.logFont)
        self.logBox.place(x=40, y=278)
        self.logBox.insert(tk.INSERT, 'Please select files \n')
        self.logBox.configure(state='disabled')
        self.logBox.tag_configure('error', foreground='red')
        self.logBox.tag_configure('warning', foreground='orange')

        # Set up the labels for displaying the selected template and lifter data file paths.
        self.selectedTemplate = Label(self, text='', bg='#FAEBD7', font=self.pathFont)
        self.selectedTemplate.place(x=230, y=78)
        self.selectedLifterData = Label(self, text='', bg='#FAEBD7', font=self.pathFont)
        self.selectedLifterData.place(x=260, y=158)

        # Set up the buttons for selecting the template and lifter data files.
        self.templateButton = Button(self, text='Select Template File', bg='#FAEBD7', font=self.buttonFont, command=self.selectTemplate)
        self.templateButton.place(x=40, y=38)
        self.lifterDataButton = Button(self, text='Select Lifter Data', bg='#FAEBD7', font=self.buttonFont, command=self.selectLifterData)
        self.lifterDataButton.place(x=40, y=118)
        self.runButton = Button(self, text='Generate Certificate', bg='#FAEBD7', font=self.buttonFont, command=self.check)
        self.runButton.place(x=40, y=198)

    def selectLifterData(self):
        """
        Opens a file dialog box to allow the user to select the lifter data Excel file.
        Sets the selected file path to the class variable self.lifterData.
        Updates the GUI to show the selected file name.
        Logs the selected file path in the log box.
        """
        # Open a file dialog box to allow the user to select the lifter data Excel file
        self.lifterData = filedialog.askopenfilename(filetypes=[('Lifter data', '*.xlsx')])
        
        # Update the GUI to show the selected file name
        self.selectedLifterData.configure(text = os.path.split(self.lifterData)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected input data:\n{self.lifterData}')


    def selectTemplate(self):
        """
        Opens a file dialog box to allow the user to select the certificate template Word file.
        Sets the selected file path to the class variable self.template.
        Updates the GUI to show the selected file name.
        Logs the selected file path in the log box.
        """
        # Open a file dialog box to allow the user to select the certificate template Word file
        self.template = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
        
        # Update the GUI to show the selected file name
        self.selectedTemplate.configure(text = os.path.split(self.template)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected template:\n{self.template}')

  
    def log(self, message, level='info'):
        """
        Logs a message to the GUI log box.

        Args:
            message (str): The message to log.
            level (str, optional): The level of the message (default 'info').

        Returns:
            None
        """
        # Enables the logBox for writing
        self.logBox.configure(state='normal')
        # Writes the message to the logBox, using the given level
        self.logBox.insert(tk.INSERT, message + "\n", level)
        # Disables the logBox after writing
        self.logBox.configure(state='disabled')
        # Scrolls to the end of the logBox
        self.logBox.see("end")
        # Updates the GUI
        self.update()


    def processData(self):
        self.log('running.....')

        wb = openpyxl.load_workbook(self.lifterData)
        ws = wb.active

        doc = Document(self.template)

        allPDFs = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] != None:
                #data tuple =  name, weight, gender, age, equipment
                #data.append((row[0],row[2],row[3],row[4],row[5]))
                self.findReplaceTable(doc,dict(zzNAMEzz=row[0],zzCLASSzz=row[2], zzGENDERzz=row[3], zzDIVzz=row[4], zzEQUIPzz=row[5]))

                self.log(f'Generating certificate for {row[0]}')

                doc.save(f'c:\\temp\\certificate_{row[0]}.docx')
                convert(f'c:\\temp\\certificate_{row[0]}.docx')
                allPDFs.append(f'c:\\temp\\certificate_{row[0]}.pdf')

                self.log(f'Generated certificate for {row[0]}')
                #revert back to save opening/close doc lots
                self.findReplaceTable(doc,dict(zzNAMEzz=row[0]),True)

        self.log(f'Merging Certificate PDFs')

        merger = PdfMerger()
        for pdf in allPDFs:
            merger.append(pdf)

        merger.write('certificates.pdf')

        self.log(f'Certificates Merged into one PDF ')



    def findReplaceTable(self, doc, data, revert=False):
        """
        Find and replace text in tables in a Word document.

        Arguments:
        - doc: a `docx.Document` object representing the Word document to modify
        - data: a dictionary containing the text to find and its replacement
        - revert: if `True`, the find and replace arguments will be swapped
            (default: False)

        Returns:
        - None

        """
        # If revert is True, swap find and replace arguments
        if revert:
            data = {v: k for k, v in data.items()}

        # Loop over the tables in the document
        for table in doc.tables:
            # Loop over the rows in the table
            for row in table.rows:
                # Loop over the cells in the row
                for cell in row.cells:
                    # If the find text is in the cell, loop over the paragraphs in the cell
                    if any(find in cell.text for find in data.keys()):
                        for p in cell.paragraphs:
                            # If the find text is in the paragraph, loop over the runs in the paragraph
                            if any(find in p.text for find in data.keys()):
                                for run in p.runs:
                                    # If the find text is in the run, replace it with the replacement text
                                    if any(find in run.text for find in data.keys()):
                                        for find, replace in data.items():
                                            run.text = run.text.replace(find, replace)



    def check(self):
        """Checks if both lifterData and template have been selected. If both have been selected, processData() is called.
        If either one or none of them has been selected, a log message is printed with a warning or error level depending on
        the number of files selected.
        """
        if self.lifterData and self.template:
            self.processData()
        elif self.lifterData:
            self.log("Please select a template file", 'warning')
        elif self.template:
            self.log("Please select lifter data", 'warning')
        else:
            self.log("Please select files", 'error')



if __name__ == '__main__':
    app = CPA_Cert_Generator()
    app.mainloop()