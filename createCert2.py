import time
import PyPDF2
import multiprocessing.dummy
import os
import random
import tkinter as tk
from tkinter import filedialog, scrolledtext, Label, Button
import docx
import docx2pdf
import openpyxl
import re
import uuid

class CPA_Cert_Generator (tk.Tk):
    def __init__(self):
        """Initialize the GUI application with buttons, labels, and a text box for logging."""

        super().__init__()

        self.lifterData =[]
        self.GUID = ""
        self.lots = {}
        self.days = 0
        self.eventName = ""

        # Set the initial values of the template and lifterData attributes to None.
        self.templateNames = {
            "certificateTemplate":"3LIFT CERT TEMPLATE.docx",
            "speakerTemplate":"SPEAKER TEMPLATE.docx",
            "weighinTemplate":"WEIGH IN TEMPLATE.docx",
            "gearcheckTemplate":"GEAR CHECK TEMPLATE.docx",
            "manualscorecardTemplate":"MANUAL SCORESHEET TEMPLATE.docx"

        }

        self.certificateTemplate = None
        self.speakerTemplate = None
        self.weighinTemplate = None
        self.gearcheckTemplate = None
        self.lifterDataInput = None
        self.manualscorecardTemplate = None

        self.certificatePDFs = []
        self.speakerPDFs = []

        # Set the size, background color, and title of the window.
        self.geometry('800x720')
        self.configure(background='#FAEBD7')
        self.title('CPA Certificate Generator')

        # Set the fonts for buttons, labels, and file path displays.
        self.buttonFont = ('courier', 12, 'normal')
        self.labelFont = ('courier', 12, 'bold')
        self.pathFont = ('courier', 12, 'normal')
        self.logFont = ('courier', 10, 'normal')

        # Set up the logging text box and configure it with tags for different message types.
        self.logBox = scrolledtext.ScrolledText(self, wrap=tk.WORD, width=85, height=10, font=self.logFont)
        self.logBox.place(x=40, y=520)
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
        
        self.eventNameField = tk.Entry(self, font=self.labelFont)
        self.eventNameField.place(x=40,y=50)
        self.eventNameLabel = create_label(self, 'eventNameLabel', 'Event Name', 40, 20)

        self.LifterDataButton = create_button(self, 'lifterDataButton', 'Select Lifter Data', 350, 60, self.selectLifterData)
        self.selectedLifterData = create_label(self, 'selectedLifterData', '<please select file>', 350, 100)

        self.certificateTemplateButton = create_button(self, 'certificateTemplateButton', 'Select Certificate Template File', 350, 130, self.selectCertificateTemplate)
        self.selectedCertificateTemplate = create_label(self, 'selectedCertificateTemplate', '<please select file>', 350, 170)

        self.speakerTemplateButton = create_button(self, 'speakerTemplateButton', 'Select Speaker Card Template File', 350, 200, self.selectSpeakerTemplate)
        self.selectedSpeakerTemplate = create_label(self, 'selectedSpeakerTemplate', '<please select file>', 350, 240)

        self.weighinTemplateButton = create_button(self, 'weighinTemplateButton', 'Select Weigh In Template File', 350, 270, self.selectWeighinTemplate)
        self.selectedWeighinTemplate = create_label(self, 'selectedWeighinTemplate', '<please select file>', 350, 310)

        self.gearcheckTemplateButton = create_button(self, 'gearcheckTemplateButton', 'Select Gear Check Template File', 350, 340, self.selectGearcheckTemplate)
        self.selectedGearcheckTemplate = create_label(self, 'selectedGearcheckTemplate', '<please select file>', 350, 380)

        self.manualscorecardTemplateButton = create_button(self, 'manualscorecardTemplateButton', 'Select Manual Score Card Template File', 350, 410, self.selectManualScoreCardTemplate)
        self.selectedManualscorecardTemplate = create_label(self, 'selectedManualscorecardTemplate', '<please select file>', 350, 450)

        self.runButton = Button(self, text='Create files', bg='#FAEBD7', font=self.buttonFont, command=self.run)
        self.runButton.place(x=40, y=480)

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

    def selectWeighinTemplate(self):
        # Open a file dialog box to allow the user to select the certificate template Word file
        self.weighinTemplate = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
        
        # Update the GUI to show the selected file name
        self.selectedWeighinTemplate.configure(text = os.path.split(self.weighinTemplate)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected template:\n{self.weighinTemplate}')

    def selectGearcheckTemplate(self):
        # Open a file dialog box to allow the user to select the certificate template Word file
        self.gearcheckTemplate = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
        
        # Update the GUI to show the selected file name
        self.selectedGearcheckTemplate.configure(text = os.path.split(self.gearcheckTemplate)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected template:\n{self.gearcheckTemplate}')

    def selectManualScoreCardTemplate(self):
        # Open a file dialog box to allow the user to select the certificate template Word file
        self.manualscorecardTemplate = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
        
        # Update the GUI to show the selected file name
        self.selectedManualscorecardTemplate.configure(text = os.path.split(self.manualscorecardTemplate)[1])
        
        # Log the selected file path in the log box
        self.log(f'selected template:\n{self.manualscorecardTemplate}')


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

        self.eventName = self.eventNameField.get()

        # Load the Excel workbook and select the active worksheet
        wb = openpyxl.load_workbook(self.lifterDataInput)
        ws = wb.active

        ID = 1

        # Iterate through each row of data in the worksheet
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Check if the first column of the row is not None
            if row[0] != None:
                day = row[0]
                flight = row[1]
                __rand = str(random.randint(0, 100000))

                if day in self.lots:
                    if flight not in self.lots[day]:
                        self.lots[day][flight] = []
                else:
                    self.lots[day] = {}
                    if flight not in self.lots[day]:
                        self.lots[day][flight] = []

                self.lots[day][flight].append({"id":ID,"rand":__rand})

                lifterDataTemplate = {
                    "Day": day,
                    "Flight": flight,
                    "Age": re.sub(r'\([^)]*\)', '', row[2]),
                    "Weight": re.sub(r'^.*\s-\s', '', row[3]),
                    "First name": row[4],
                    "Last name": row[5],
                    "Gender": row[6],
                    "DOB": row[7],
                    "Nation": "New Zealand",
                    "Association": row[8],
                    "Raw": row[9],
                    "Instagram": row[10],
                    "Notes": row[11],
                    "ID": ID,
                    "Lot": "0"
                }

                self.lifterData.append(lifterDataTemplate)

                ID += 1

        #set lot numbers
        for day, flight in self.lots.items():
            for flight, v in flight.items():
                lotsSorted = sorted(v, key=lambda x: x['rand'])
                for i in range(0, len(lotsSorted)):
                    self.lifterData[lotsSorted[i]["id"]-1]["Lot"]=str(i+1)
        
        #get number of days
        self.days = max(list(set(item["Day"] for item in self.lifterData)))

    def findReplaceParagraph(self, doc, data, revert=False):
        # If revert is True, swap find and replace arguments
        if revert:
            data = {v: k for k, v in data.items()}

        for paragraph in doc.paragraphs:
            for key, value in data.items():
                if key in paragraph.text:
                    for run in paragraph.runs:
                        if key in run.text:
                            run.text = run.text.replace(key, value)

    def createGearCheck(self):
        allPDFs = []
        self.log("Creating gear check sheets")
        for i in range(0,self.days):
            i += 1
            curDay = [d for d in self.lifterData if d['Day'] == i]

            self.log(f"Gear check day {i}")

            doc = docx.Document(self.gearcheckTemplate)
            table = doc.tables[0]
            for j in sorted(curDay, key=lambda x: (x['Flight'],int(x['Lot']))):
                rc = table.add_row().cells
                rc[0].text = f"{j['First name']} {j['Last name']}"  
        
            
            self.findReplaceParagraph(doc,dict(zzEVENTzz=self.eventName,zzSESSIONzz=str(i)))

            time.sleep(0.5)
            doc.save(f'c:\\temp\\{self.GUID}\\gear_check{str(i)}.docx')
            time.sleep(0.5)

            docx2pdf.convert(f'c:\\temp\\{self.GUID}\\gear_check{str(i)}.docx')

            # Add the path of the generated PDF to the list
            allPDFs.append(f'c:\\temp\\{self.GUID}\\gear_check{str(i)}.pdf')
        
        

        # Create a new PDF merger object
        merger = PyPDF2.PdfMerger()

        # Iterate through each PDF file path in the list and add it to the merger
        for pdf in allPDFs:
            merger.append(pdf)

        # Write the merged PDF to a file
        merger.write('gear check.pdf')
        self.log("Gear check created")

    def createWeighIn(self):
        allPDFs = []
        self.log("Creating weigh ins")
        for i in range(0,self.days):
            i += 1
            curDay = [d for d in self.lifterData if d['Day'] == i]

            self.log(f"Weigh ins day {i}")

            doc = docx.Document(self.weighinTemplate)
            table = doc.tables[0]
            for j in sorted(curDay, key=lambda x: (x['Flight'],int(x['Lot']))):
                rc = table.add_row().cells
                rc[0].text = str(j['Day'])
                rc[1].text = j["Flight"]
                rc[2].text = f"{j['First name']} {j['Last name']}"
                rc[3].text = j['Lot']            
            
            time.sleep(0.5)
            doc.save(f'c:\\temp\\{self.GUID}\\weigh_in_day_{str(i)}.docx')
            time.sleep(0.5)

            docx2pdf.convert(f'c:\\temp\\{self.GUID}\\weigh_in_day_{str(i)}.docx')

            # Add the path of the generated PDF to the list
            allPDFs.append(f'c:\\temp\\{self.GUID}\\weigh_in_day_{str(i)}.pdf')

        
        
        # Create a new PDF merger object
        merger = PyPDF2.PdfMerger()

        # Iterate through each PDF file path in the list and add it to the merger
        for pdf in allPDFs:
            merger.append(pdf)

        # Write the merged PDF to a file
        merger.write('weigh in.pdf')
        self.log("Weigh ins created")
        
    def runLifterSpecific(self):

        self.log(f"Generating certificates and speaker cards\nThis will take a couple of minutes")
        pool1= multiprocessing.Pool(processes=multiprocessing.cpu_count())

        pdfgen = PDFGenerator()

        args =  [(lifter, self.certificateTemplate, self.GUID) for lifter in self.lifterData]
        resultsCert = pool1.map(pdfgen.createCetificates, args)
        args =  [(lifter, self.speakerTemplate, self.GUID) for lifter in self.lifterData]
        resultsSpeaker = pool1.map(pdfgen.createSpeaker, args)

        # close the pool and wait for the work to finish
        pool1.close()
        pool1.join()

        #docx2pdf.convert all docx in temp dir to pdfs
        docx2pdf.convert(f'c:\\temp\\{self.GUID}')


        merger1 = PyPDF2.PdfMerger()
        # Iterate through each PDF file path in the list and add it to the merger
        for pdf in resultsCert:
            merger1.append(pdf.replace('.docx','.pdf'))
        # Write the merged PDF to a file
        merger1.write('certificates.pdf')


        merger2 = PyPDF2.PdfMerger()
        # Iterate through each PDF file path in the list and add it to the merger
        for pdf in resultsSpeaker:
            merger2.append(pdf.replace('.docx','.pdf'))
        # Write the merged PDF to a file
        merger2.write('speaker cards.pdf')

        self.log(f"Finished generating certificates and speaker cards")

    def createManualScoreSheet(self):
            allPDFs = []
            self.log("Creating manual scoresheets")
            for i in range(0,self.days):
                i += 1
                curDay = [d for d in self.lifterData if d['Day'] == i]

                self.log(f"Scoresheet day {i}")

                doc = docx.Document(self.manualscorecardTemplate)
                table = doc.tables[0]
                for j in sorted(curDay, key=lambda x: (x['Flight'],int(x['Lot']))):
                    rc = table.add_row().cells
                    rc[0].text = f"{j['First name']} {j['Last name']}"  
                    rc[1].text = f"" #teamS
                    rc[2].text = f"{j['Age']}"
                    rc[3].text = f"{j['Lot']}"
                    rc[5].text = f"{j['Weight']}"
            
                
                self.findReplaceParagraph(doc,dict(zzEVENTzz=self.eventName,zzSESSIONzz=str(i)))

                time.sleep(0.5)
                doc.save(f'c:\\temp\\{self.GUID}\\manual_scoresheet{str(i)}.docx')
                time.sleep(0.5)

                docx2pdf.convert(f'c:\\temp\\{self.GUID}\\manual_scoresheet{str(i)}.docx')

                # Add the path of the generated PDF to the list
                allPDFs.append(f'c:\\temp\\{self.GUID}\\manual_scoresheet{str(i)}.pdf')
            
            

            # Create a new PDF merger object
            merger = PyPDF2.PdfMerger()

            # Iterate through each PDF file path in the list and add it to the merger
            for pdf in allPDFs:
                merger.append(pdf)

            # Write the merged PDF to a file
            merger.write('manual scoresheet.pdf')
            self.log("Scoresheet created")


    def run(self):
        self.proccessData()
        self.GUID = str(uuid.uuid4())
        os.mkdir(f'c:\\temp\\{self.GUID}\\')

        self.runLifterSpecific()
        self.createWeighIn()
        self.createGearCheck()
        self.createManualScoreSheet()

        self.log("~~~~~~~~~~~~~~~~~\nCompleted!\n~~~~~~~~~~~~~~~~~")

class PDFGenerator:
    def __init__(self):
        self.speakerPDFs = []
        self.certificatePDFs= []
    def createSpeaker(self, data):
        lifter, speakerTemplate, GUID = data
        # Create a new Word docx.document based on the provided template
        doc = docx.Document(speakerTemplate)


        name = f'{lifter["First name"]} {lifter["Last name"]}'
        dob = lifter["DOB"].strftime("%d/%m/%Y")
        
        #spaces are so it doesn't replace other numbers in the template
        self.findReplaceTable(doc,dict(zzNAMEzz=name,zzCLASSzz=lifter["Weight"], zzDOBzz=dob, zzNATIONzz=lifter["Nation"], zzLOTzz=lifter["Lot"]+"   "))

        time.sleep(0.1)
        doc.save(f'c:\\temp\\{GUID}\\speaker_{name}.docx')
        time.sleep(0.1)
        return(f'c:\\temp\\{GUID}\\speaker_{name}.docx')

    def createCetificates(self, data):

        lifter, certificateTemplate, GUID = data

        doc = docx.Document(certificateTemplate)

        name = f'{lifter["First name"]} {lifter["Last name"]}'

        self.findReplaceTable(doc,dict(zzNAMEzz=name,zzCLASSzz=lifter["Weight"], zzGENDERzz=lifter["Gender"], zzDIVzz=lifter["Age"], zzEQUIPzz=lifter["Raw"]))

        time.sleep(0.1)
        doc.save(f'c:\\temp\\{GUID}\\certificate_{name}.docx')
        time.sleep(0.1)
        return (f'c:\\temp\\{GUID}\\certificate_{name}.docx')

    def findReplaceTable(self, doc, data, revert=False):
        """
        Find and replace text in tables in a Word docx.document.

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
        




if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = CPA_Cert_Generator()
    app.mainloop()
    print()