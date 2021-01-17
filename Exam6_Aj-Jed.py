import os
from datetime import date
from openpyxl import *
from pathlib import Path
import zipfile

#กำหนดวันที่ปัจจุบัน
today = date.today()
d1 = today.strftime("%d%m%Y")

########### Function #########################
#Function สำหรับตรวจสอบ file size ในแต่ละ Directory
def sum_directory_size(x):
    root_directory = Path(x)
    sum_size = sum(f.stat().st_size for f in root_directory.glob('**/*') if f.is_file())
    return sum_size/1000

#Function สำหรับตรวจสอบ file size ในแต่ละ Zip Directory
def sum_directory_size_zip(x):
    root_directory = Path(x)
    sum_size = sum(f.stat().st_size for f in root_directory.glob('**/*.zip') if f.is_file())
    return sum_size/1000

#Function สำหรับ Zip Directory
def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), os.path.join(path, '..')))
            
#Main Function สำหรับ Zip Directory           
def sizing_report(path):
    #Convert directory to zip
    for i in os.listdir(path):
        zipf = zipfile.ZipFile(path+str(i)+".zip", 'w', zipfile.ZIP_DEFLATED)
        zipdir(path+str(i), zipf)
        zipf.close() 
        
    # กำหนดตัวแปร Path ที่ต้องการทำงาน
    source_path = path
    
    # กำหนดตัวแปรผลรวม size directory ทั้งก่อนและหลัง Zip พร้อมตัวแปร Compare ratio
    before = sum_directory_size(source_path)
    after =  sum_directory_size_zip(source_path) 
    compare_ratio = -((before - after)/before) *100
    
    #Create a new Workbook
    wb = Workbook()
    #Active sheet
    sheet =  wb.active
    #Header
    sheet['A1'].value = "ZIP file_compare Report as at "  + str(d1)
    sheet['A3'].value = "Source path = " + str(source_path)
    sheet['A8'].value = "Directory Name"
    sheet['B8'].value = "Directory_before_zip size"
    #ตัวแปรสำหรับเปรียบเทียบข้อมูล
    sheet['A4'].value = "Directory size before zip : "  +str(before) + " Kbyte"
    sheet['A5'].value = "Directory size after zip : "  +str(after) + " Kbyte"
    sheet['A6'].value = "compare ratio : "  +str(compare_ratio) +" %"
    
    #Select starting row and column
    col = "A" #Select column A
    row = 9 #Select column A8 to write Directory name
    
    col2 = "B" #Select column B
    row2 = 9 #Select column B8 to write Directory size
    
    #Find all Sub_directoty
    for i in os.listdir(source_path):
        if not i.endswith(".zip"):
            sheet['{0}{1}'.format(col, row)].value = "Directory_name : " +str(i) #Directory name starting in Row A8
            row += 1 #Next rows
            sheet['{0}{1}'.format(col2, row2)].value = str(sum_directory_size(str(source_path)+str(i))) + " Kbyte." #Directory size before zip starting in Row B8
            row2 += 1 #Next row   

    #Save files
    wb.save(source_path+'Compare_Sizing_Report_' +str(d1)+'.xlsx')