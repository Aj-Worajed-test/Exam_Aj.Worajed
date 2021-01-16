from pathlib import Path
from zipfile import ZipFile
import os
from os.path import basename
import shutil
from datetime import date
#import openpyxl
from openpyxl import *
import datetime
import stat
import time

#กำหนดวันที่ปัจจุบัน
today = date.today()
d1 = today.strftime("%d%m%Y")

#Function สำหรับเรียก Modify date
def modification_date(filename):
    t = os.path.getmtime(filename)
    return datetime.datetime.fromtimestamp(t)

#Function สำหรับตรวจสอบ file size ในแต่ละ Directory
def sum_directory_size(x):
    root_directory = Path(x)
    sum_size = sum(f.stat().st_size for f in root_directory.glob('**/*') if f.is_file())
    return sum_size/1000

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
 
source_path = "C:/Users/veerachai/Desktop/test_Aj_Jed/"
sheet['A3'].value = "Source path = " + str(source_path)

#Find all Sub_directoty
for i in os.listdir(source_path):     
    sheet['{0}{1}'.format(col, row)].value = "Directory_name : " +str(i) #Directory name starting in Row A8
    row += 1 #Next rows
    
    sheet['{0}{1}'.format(col2, row2)].value = str(sum_directory_size(str(source_path)+str(i))) + " Kbyte." #Directory size starting in Row B8
    row2 += 1 #Next row
    
    #Save files
    wb.save(source_path+'Sizing_Report_' +str(d1)+'.xlsx')