'''
便民工程 查分脚本
第一步 pip install openpyxl
第二步 选择相应课程代码
第三步 输入UniqueID
第四步 F5
'''
#Check Your Grade (ICS46/33)
fileURL46 = "https://www.ics.uci.edu/~pattis/ICS-46/ics46fal19grades.zip"
fileURL33 = "https://www.ics.uci.edu/~pattis/ICS-33/ics33fal19grades.zip"

fileURL = fileURL46 #Select Course Here

ID = "123456" #Enter Your HashID Here

import zipfile
import urllib.request
import os
import openpyxl as xl
import ssl

print("Downloading...")

urllib.request.urlretrieve(fileURL, "temp.zip")

print("Extracting...")

with zipfile.ZipFile("temp.zip",'r') as zip_ref:
    zip_ref.extractall()

print("Reading...")

book = xl.load_workbook(filename = f"{'ics46fal19grades.xlsm' if '46' in fileURL else 'ics33fal19grades.xlsm'}", data_only = True)

sheet_ranges = book["Fall 2019"]

allID = sheet_ranges['A']

row = 0
for i in allID:
    if (i.value == int(ID)):
        row = i.row

for i in sheet_ranges[1][:9]:
    print (str(i.value)[:6] ,"\t",end = '')
print()
for i in sheet_ranges[row][:9]:
    print (str(i.value)[:6] ,"\t",end = '')
print()

for i in sheet_ranges[1][9:20]:
    print (str(i.value[:6]) ,"\t",end = '')
print()
for i in sheet_ranges[row][9:20]:
    print (str(i.value)[:6] ,"\t",end = '')
print()
for i in sheet_ranges[1][24:28]:
    print (str(i.value)[:6] ,"\t",end = '')
print()
for i in sheet_ranges[row][24:28]:
    print (str(i.value)[:6] ,"\t",end = '')
print()

print("Clearing...")

os.remove(f"{'ics46fal19grades.xlsm' if '46' in fileURL else 'ics33fal19grades.xlsm'}")
os.remove("temp.zip")
