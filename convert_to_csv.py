import os
import pandas as pd
from tkinter import filedialog
import re
import openpyxl
from tkinter import *
from docx2pdf import convert
def browseFiles():
    source_file = filedialog.askdirectory(master='', initialdir='', title="Select your Source directory", mustexist=True)
    return source_file

def convert_excel_to_csv(file_path, file_extension):
    cleanedname = os.path.splitext(os.path.basename(file_path))[0]
    xlfile = pd.ExcelFile(file_path)
    sheets = xlfile.sheet_names
    
    for eachsheet in sheets:
        sheet = xlfile.parse(eachsheet)
        if sheet.empty:
            # If the sheet is empty, you can skip conversion or perform additional actions
            print(f"Sheet '{eachsheet}' is empty. Skipping conversion or performing additional actions.")
            continue
        
        csvname = f"{cleanedname}-{eachsheet}.csv" if len(sheets) > 1 else f"{cleanedname}.csv"
        sheet.to_csv(csvname, index=False)
        print(csvname)
def convert_docx_to_pdf():
    path = browseFiles()
    files = os.listdir(path)
    convert(path, output_path=".")
    open_pdf_files(".")

def Convert():
    path = browseFiles()
    files = os.listdir(path)

    for eachfile in files:
        file_path = os.path.join(path, eachfile)

        if eachfile.endswith((".xlsx", ".xls", ".ods")):
            convert_excel_to_csv(file_path, os.path.splitext(eachfile)[1])
    open_csv_files(".")
# def Delete():
#     path = os.getcwd()
#     files = os.listdir(path)
    
#     for eachfile in files:
#         if eachfile.endswith("delete me.csv"):
#             os.remove(eachfile)
def open_csv_files(directory):
    csv_files = [f for f in os.listdir(directory) if f.endswith('.csv')]
    for csv_file in csv_files:
        os.startfile(os.path.join(directory, csv_file))
def open_pdf_files(directory):
    pdf_files = [f for f in os.listdir(directory) if f.endswith('.pdf')]
    for csv_file in pdf_files:
        os.startfile(os.path.join(directory, csv_file))

root = Tk()
main_container = Frame(root)
main_container.grid()
myButton = Button(main_container, text="Select Directory To Check for excel files",
                  command=Convert)
myButton1 = Button(main_container, text="convert docx to pdf",
                  command=convert_docx_to_pdf)


myButton.grid(column=0, row=0, padx=5 ,pady=5, sticky='w')
myButton1.grid(column=0, row=1, padx=5 ,pady=5, sticky='w')
# myButton3.grid(column=0, row=2, padx=5 ,pady=5, sticky='w')
root.mainloop()