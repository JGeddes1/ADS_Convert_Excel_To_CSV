import os
import pandas as pd
from tkinter import filedialog
import re
import openpyxl
from tkinter import *


# --BUTTONS START--
def browseFiles():
    source_file = filedialog.askdirectory(master='', initialdir='', title="Select your Source directory", mustexist=True)
    return source_file

def Convert():
    
    path = browseFiles()
    # path = filedialog.askdirectory(master='', initialdir='', title="Select your Source directory", mustexist=True)
    files = os.listdir(path)



    for eachfile in files:
        if eachfile.endswith(".xlsx"):
            # wb = pd.load_workbook(filename=eachfile, read_only=False)
            cleanedname = eachfile.replace(".xlsx", "")
            xlfile = pd.ExcelFile(os.path.join(path, eachfile))
            sheets_in = xlfile.book.worksheets
        

            empty_sheets = []
            for sheet in sheets_in:
                
                print(sheet.title, sheet.sheet_state, sheet.max_column)
                if sheet.max_column ==1:
                    empty_sheets.append(sheet)

                    # del eachfile['Sheet3sadsadf']

            sheets = xlfile.sheet_names
        
            print("sheets name" + eachfile + " "+ str(sheets))
            
            if len(sheets) > 1:
                
                for eachsheet in sheets:
                        
                        sheetdata = xlfile.parse(eachsheet)
                        # sheetdata.to_csv()
                        
                        csvname = cleanedname + "-" + eachsheet + ".csv"
                        if csvname == cleanedname + "-" + "ADS_spreadsheet_metadata" + ".csv" :
                            csvname = cleanedname + ".csv"
                        if csvname == cleanedname + "-" + "Dropdown" + ".csv" :
                            csvname = cleanedname + "delete me"+ ".csv"

                        x = re.search("-Sheet1", csvname)
                        xy = re.search("raster", csvname)
                        xyz = re.search("-Sheet3", csvname)
                        if x and xy:
                            csvname = cleanedname + ".csv" 
                        if xyz and xy:
                            csvname = cleanedname + "delete me"+ ".csv"
                            
                        # if 'raster_metadata-Sheet1' in csvname:
                        #     csvname = cleanedname + ".csv"  
                        # if 'raster_metadata-Sheet3' in csvname:
                        #     csvname = cleanedname + "delete me"+ ".csv"  
                            
                        sheetdata.to_csv(csvname, index=False)
                        print(csvname)
            else:
                for eachsheet in sheets:
                    sheetdata = xlfile.parse(eachsheet)
                    # sheetdata.to_csv()
            
                    csvname = cleanedname + ".csv"
                    sheetdata.to_csv(csvname, index=False)
                    print(csvname)
        if eachfile.endswith(".delete me.csv"):
            os.remove(eachfile)


        if eachfile.endswith(".xls"):
            cleanedname = eachfile.replace(".xls", "")
            xlfile = pd.ExcelFile(os.path.join(path, eachfile))
            sheets = xlfile.sheet_names
            print("sheets listed" + str(sheets))
            if len(sheets) > 1: 
                for eachsheet in sheets:
                    sheetdata = xlfile.parse(eachsheet)
                        # sheetdata.to_csv()
                
                    csvname = cleanedname + "-" + eachsheet + ".csv"
                            # csvname = cleanedname + ".csv"
                            # print(csvname +'help me')
                    sheetdata.to_csv(csvname, index=False)
                    print(csvname)
            else:
                for eachsheet in sheets:
                    sheetdata = xlfile.parse(eachsheet)
                    # sheetdata.to_csv()
                
                    csvname = cleanedname + ".csv"
                    sheetdata.to_csv(csvname, index=False)
                    print(csvname)
            


        if eachfile.endswith(".ods"):
            cleanedname = eachfile.replace(".ods", "")
            xlfile = pd.ExcelFile(os.path.join(path, eachfile))
            sheets = xlfile.sheet_names
            print("sheets listed" + str(sheets))
            if len(sheets) > 1: 
                for eachsheet in sheets:
                    sheetdata = xlfile.parse(eachsheet)
                        # sheetdata.to_csv()
                
                    csvname = cleanedname + "-" + eachsheet + ".csv"
                    sheetdata.to_csv(csvname, index=False)
                    print(csvname)
            else:
                for eachsheet in sheets:
                    sheetdata = xlfile.parse(eachsheet)
                    # sheetdata.to_csv()
                
                    csvname = cleanedname + ".csv"
                    sheetdata.to_csv(csvname, index=False)
                    print(csvname)


def Delete():
    
    path = os.getcwd()
    files = os.listdir(path)
    
    for eachfile in files:
        if eachfile.endswith("delete me.csv"):
            os.remove(eachfile)

def Boo():
    label = Label(main_container, text="Boo! I got you Evelyn :)")
    label.grid(column=0, row=3, padx=5 ,pady=5, sticky='w')

root = Tk()
main_container = Frame(root)
main_container.grid()
myButton = Button(main_container, text="Select Directory To Check for excel files",
                  command=Convert)
myButton1 = Button(main_container, text="Delete Files that are produced that say delete me",
                  command=Delete)

myButton3 = Button(main_container, text="Press me for suprise",
                  command=Boo)

myButton.grid(column=0, row=0, padx=5 ,pady=5, sticky='w')
myButton1.grid(column=0, row=1, padx=5 ,pady=5, sticky='w')
myButton3.grid(column=0, row=2, padx=5 ,pady=5, sticky='w')
root.mainloop()

