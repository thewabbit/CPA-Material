import os
import random
import tkinter as tk
from tkinter import filedialog, scrolledtext, Label, Button
from docx import Document
from PyPDF2 import PdfMerger
from docx2pdf import convert
import openpyxl
import re
import uuid

class CPA_Cert_Generator (tk.Tk):
    def __init__(self):
        """Initialize the GUI application with buttons, labels, and a text box for logging."""

        super().__init__()

        self.lifterData = []
        self.GUID = ""

        # Set the initial values of the template and lifterData attributes to None.
        self.templateNames = {
            "certificateTemplate":"3LIFT CERT TEMPLATE.docx",
            "speakerTemplate":"SPEAKER TEMPLATE.docx"
        }

        self.certificateTemplate = None
        self.speakerTemplate = None
        self.lifterDataInput = None

        # testing
        # self.certificateTemplate = r'G:\My Drive\Documents\Powerlifting\CPA\Create Certificates\Templates\3LIFT CERT TEMPLATE.docx'
        # self.lifterDataInput = r'G:\My Drive\Documents\Powerlifting\CPA\Create Certificates\Inputs\Copy of Provincials 2023 final nominations2.xlsx'

        # Set the size, background color, and title of the window.
        self.geometry('580x580')
        self.configure(background='#FAEBD7')
        self.title('CPA Certificate Generator')

        # Set the fonts for buttons, labels, and file path displays.
        self.buttonFont = ('courier', 12, 'normal')
        self.labelFont = ('courier', 12, 'bold')
        self.pathFont = ('courier', 12, 'normal')
        self.logFont = ('courier', 10, 'normal')

        # Set up the logging text box and configure it with tags for different message types.
        self.logBox = scrolledtext.ScrolledText(self, wrap=tk.WORD, width=60, height=10, font=self.logFont)
        self.logBox.place(x=40, y=318)
        self.logBox.insert(tk.INSERT, 'Please select files \n')
        self.logBox.configure(state='disabled')
        self.logBox.tag_configure('error', foreground='red')
        self.logBox.tag_configure('warning', foreground='orange')


        def create_button(parent, name, text, x, y, command):
            button = Button(parent, name=name, text=text, bg='#FAEBD7', font=self.buttonFont, command=command)
            button.place(x=x, y=y)
            return button

        def create_label(parent, name, text, x, y):
            label = Label(parent, name=name, text=text, bg='#FAEBD7', font=self.pathFont)
            label.place(x=x, y=y)
            return label

        self.LifterDataButton = create_button(self, 'lifterDataButton', 'Select Lifter Data', 40, 38, self.selectLifterData)
        self.selectedLifterData = create_label(self, 'selectedLifterData', '', 40, 78)

        self.certificateTemplateButton = create_button(self, 'certificateTemplateButton', 'Select Certificate Template File', 40, 118, self.selectCertificateTemplate)
        self.selectedCertificateTemplate = create_label(self, 'selectedCertificateTemplate', '', 40, 158)

        self.speakerTemplateButton = create_button(self, 'speakerTemplateButton', 'Select Speaker Card Template File', 40, 198, self.selectSpeakerTemplate)
        self.selectedSpeakerTemplate = create_label(self, 'selectedSpeakerTemplate', '', 40, 238)


        self.runButton = Button(self, text='Create files', bg='#FAEBD7', font=self.buttonFont, command=self.run)
        self.runButton.place(x=40, y=278)

        self.templateScan()

    def templateScan(self):
        for k,v in self.templateNames.items():
            path = os.path.join(os.getcwd(),"Templates",v)
            if os.path.exists(path):
                setattr(self,k,path)
                v1 = f"selected{k[0].upper() + k[1:]}"
                getattr(self, v1).configure(text = v)


    def selectLifterData(self):
        # Open a file dialog box to allow the user to select the lifter data Excel file
        self.lifterDataInput = filedialog.askopenfilename(filetypes=[('Lifter data', '*.xlsx')])
        
        # Update the GUI to show the selected file name
        self.selectedLifterData.configure(text = os.path.split(self.lifterDataInput)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected input data:\n{self.lifterDataInput}')

    def selectCertificateTemplate(self):
        # Open a file dialog box to allow the user to select the certificate template Word file
        self.certificateTemplate = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
        
        # Update the GUI to show the selected file name
        self.selectedCertificateTemplate.configure(text = os.path.split(self.certificateTemplate)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected template:\n{self.certificateTemplate}')

    def selectSpeakerTemplate(self):
        # Open a file dialog box to allow the user to select the certificate template Word file
        self.speakerTemplate = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
        
        # Update the GUI to show the selected file name
        self.selectedSpeakerTemplate.configure(text = os.path.split(self.speakerTemplate)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected template:\n{self.speakerTemplate}')

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

    def proccessData(self):
        self.log('compiling lifter input data')

        # Load the Excel workbook and select the active worksheet
        wb = openpyxl.load_workbook(self.lifterDataInput)
        ws = wb.active

        # Iterate through each row of data in the worksheet
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Check if the first column of the row is not None
            if row[0] != None:
                lifterDataTemplate = {
                    "Day": row[0],
                    "Flight": row[1],
                    "Age": re.sub(r'\([^)]*\)', '', row[2]),
                    "Weight": re.sub(r'^.*\s-\s', '', row[3]),
                    "First name": row[4],
                    "Last name": row[5],
                    "Gender": row[6],
                    "Date Of Birth": row[7],
                    "Association": row[8],
                    "Raw": row[9],
                    "Instagram": row[10],
                    "Notes": row[11],
                    "ID": random.randint(0, 100000)
                }

                self.lifterData.append(lifterDataTemplate)

    def createCetificates(self):
        """
        Processes the data in the Excel workbook and generates certificates for each row of data.
        Merges all generated certificates into a single PDF file.

        Returns:
        None
        """

        # Create a new Word document based on the provided template
        doc = Document(self.certificateTemplate)

        # Initialize a list to store the paths of all generated PDFs
        allPDFs = []

        # Iterate through each row of data in the worksheet
        for l in self.lifterData:

            name = f'{l["First name"]} {l["Last name"]}'

            self.findReplaceTable(doc,dict(zzNAMEzz=name,zzCLASSzz=l["Weight"], zzGENDERzz=l["Gender"], zzDIVzz=l["Age"], zzEQUIPzz=l["Raw"]))

            # Log that a certificate is being generated for the current row
            self.log(f'Generating certificate for {name}')

            # Save the current document as a Word file and convert it to PDF
            doc.save(f'c:\\temp\\{self.GUID}\\certificate_{name}.docx')
            convert(f'c:\\temp\\{self.GUID}\\certificate_{name}.docx')
            # Add the path of the generated PDF to the list
            allPDFs.append(f'c:\\temp\\{self.GUID}\\certificate_{name}.pdf')

            # Log that a certificate has been generated for the current row
            self.log(f'Generated certificate for {name}')

            # Replace placeholder values in the Word document with empty strings to revert back to the original template
            self.findReplaceTable(doc,dict(zzNAMEzz=name,zzCLASSzz=l["Weight"], zzGENDERzz=l["Gender"], zzDIVzz=l["Age"], zzEQUIPzz=l["Raw"]), True)

        # Log that the certificates are being merged into a single PDF
        self.log(f'Merging Certificate PDFs')

        # Create a new PDF merger object
        merger = PdfMerger()

        # Iterate through each PDF file path in the list and add it to the merger
        for pdf in allPDFs:
            merger.append(pdf)

        # Write the merged PDF to a file
        merger.write('certificates.pdf')

        # Log that the certificates have been merged into a single PDF
        self.log(f'Certificates merged into one PDF')

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

    def run(self):
        self.proccessData()
        self.GUID = str(uuid.uuid4())
        os.mkdir(f'c:\\temp\\{self.GUID}\\')
        self.createCetificates()
        #self.createWeighIn()


if __name__ == '__main__':
    app = CPA_Cert_Generator()
    app.mainloop()