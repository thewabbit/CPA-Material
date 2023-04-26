import openpyxl
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfFileMerger
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import ttk
from tkinter import * 

# Step 1: Create a function to handle button click event
def process_data():
    # Step 2: Open file dialog to select Excel file
    root = tk.Tk()
    root.withdraw()
    inputdata = filedialog.askopenfilename(filetypes=[('Lifter data', '*.xlsx')])
    template = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])


    # Step 3: Load Excel file and read in names
    wb = openpyxl.load_workbook(inputdata)
    ws = wb.active

    names = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        names.append(row[0])

    # Step 4: Load Word certificate template
    doc = Document(template)

    # Step 5: Replace name placeholder with each name in Excel file
    for name in names:
        for paragraph in doc.paragraphs:
            if '<<name>>' in paragraph.text:
                paragraph.text = paragraph.text.replace('<<name>>', name)

        # Step 6: Save certificate as PDF file
        doc.save(f'certificate_{name}.docx')
        convert(f'certificate_{name}.docx')

    # Step 7: Merge all PDF files into one document
    pdfs = [f'certificate_{name}.pdf' for name in names]

    merger = PdfFileMerger()
    for pdf in pdfs:
        merger.append(pdf)

    merger.write('certificates.pdf')
    tk.messagebox.showinfo(title='Process Complete', message='Certificates have been generated successfully.')

def process_data():
    print("proces")

def selectLifterData(selectedLifterData,logBox):
    root = tk.Tk()
    root.withdraw()
    inputdata = filedialog.askopenfilename(filetypes=[('Lifter data', '*.xlsx')])
    selectedLifterData.configure(text = inputdata)
    logBox.configure(text = inputdata)

    
def selectTemplate(selectedTemplate,logBox):
    root = tk.Tk()
    root.withdraw()
    template = filedialog.askopenfilename(filetypes=[('Template file', '*.docx')])
    selectedTemplate.configure(text = template)
    logBox.configure(state='normal')
    logBox.insert(tk.INSERT, 'selected template:\n{0}\n'.format(template))
    logBox.configure(state='disabled')
    logBox.see("end")


def createWindow():
    root = Tk()
    # create the main window
    root.geometry('580x470')
    root.configure(background='#FAEBD7')
    root.title('CPA Certificate Generator')

    # set fonts
    buttonFont = ('courier', 12, 'normal')
    labelFont = ('courier', 12, 'bold')
    pathFont = ('courier', 12, 'normal')
    logFont = ('courier', 10, 'normal')

    # log
    Label(root, text='Log:', bg='#FAEBD7', font=labelFont).place(x=40, y=248)
    logBox = scrolledtext.ScrolledText(root, wrap = tk.WORD, width = 60, height = 10, font = logFont)
    logBox.place(x=40, y=278)
    logBox.insert(tk.INSERT, 'Please select files \n')
    logBox.configure(state='disabled')
    

    # create template button
    templateButton = Button(root, text='Select Template File', bg='#FAEBD7', font=buttonFont, command=lambda:selectTemplate(selectedTemplate,logBox))
    templateButton.place(x=40, y=38)

    Label(root, text='Selected Template:', bg='#FAEBD7', font=labelFont).place(x=40, y=78)

    selectedTemplate = Label(root, text='', bg='#FAEBD7', font=pathFont)
    selectedTemplate.place(x=230, y=78)


    # create lifter button
    lifterDataButton = Button(root, text='Select Lifter Data', bg='#FAEBD7', font=buttonFont, command=lambda:selectLifterData(selectedLifterData,logBox))
    lifterDataButton.place(x=40, y=118)

    Label(root, text='Selected Lifter Data:', bg='#FAEBD7', font=labelFont).place(x=40, y=158)

    selectedLifterData = Label(root, text='', bg='#FAEBD7', font=pathFont)
    selectedLifterData.place(x=260, y=158)


    # run process
    runButton = Button(root, text='Generate Certificate', bg='#FAEBD7', font=buttonFont, command=process_data)
    runButton.place(x=40, y=198)

    


    root.mainloop()



if __name__ == "__main__":
    createWindow()


