# Exam5_Sizing directory Report

# Requirements
- Python 3.6+
- pip install openpyxl
- pip install pathlib

## Installation steps
- pip install os
- pip install openpyxl
- pip install pathlib
- pip install datetime

## How to use and Examples
from pathlib import Path
import os
from datetime import date
from openpyxl import *

#Set Date format
today = date.today()
d1 = today.strftime("%d%m%Y")

#Function for create Modify date
def modification_date(filename):
    t = os.path.getmtime(filename)
    return datetime.datetime.fromtimestamp(t)

#Function for calculate file size in each Directory
def sum_directory_size(x):
    root_directory = Path(x)
    sum_size = sum(f.stat().st_size for f in root_directory.glob('**/*') if f.is_file())
    return sum_size/1000 #for Kbyte

#For generate Report
def sizing_report(path):
    #Create a new Workbook
    wb = Workbook()
    
    #Active sheet
    sheet =  wb.active
    
    #Header
    sheet['A1'].value = "Sizing Report as at "  + str(d1)
    sheet['A5'].value = "Directory Name"
    sheet['B5'].value = "Directory size"
    
    #Select starting row and column
    col = "A" #Select column A
    row = 6 #Select column A8 to write Directory name
    
    col2 = "B" #Select column B
    row2 = 6 #Select column B8 to write Directory size
     
    source_path = path
    sheet['A3'].value = "Source path = " + str(source_path)
    
    #Find all Sub_directoty
    for i in os.listdir(source_path):     
        sheet['{0}{1}'.format(col, row)].value = "Directory_name : " +str(i) #Directory name starting in Row A8
        row += 1 #Next rows
        
        sheet['{0}{1}'.format(col2, row2)].value = str(sum_directory_size(str(source_path)+str(i))) + " Kbyte." #Directory size starting in Row B8
        row2 += 1 #Next row
        
        #Save files
        wb.save(source_path+'Sizing_Report_' +str(d1)+'.xlsx')
        
##Generate Raport
sizing_report("C:/Users/my_name/Desktop/exe5/My_folder/")


```

# Contributor
Veerachai Mitmorn (Vee)




