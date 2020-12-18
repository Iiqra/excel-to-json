import xlrd
from collections import OrderedDict
import json

dbFilename = 'studentDB.xlsx'
tblStudent = 0
fldsStudent = 'Id,Name,Age,GPA' 
 

def readExcelSheet(filename, index):
     wb = xlrd.open_workbook(filename)
     sh = wb.sheet_by_index(index)
     return sh

def mapDataFields(sheet, fields):
     data_list = []
     fieldList = fields.split(',')
     print (fieldList) # id, name, age, GPA 
     for rownum in range(1, sheet.nrows):
          data = OrderedDict()
          i = 0 
          row_values = sheet.row_values(rownum)
          print (row_values)
          while(i < len(fieldList)):
               data[fieldList[i]] = row_values[i]
               print (data) 
               i +=1
          data_list.append(data)
     print(data_list)
     return data_list 

def writeToJsonFile(fileName, dataList):
     with open(fileName, "w", encoding="utf-8") as writeJsonfile:
           json.dump(dataList, writeJsonfile, indent=4,default=str)

print('*** Reading Data from Excel Sheet***')
sh = readExcelSheet(dbFilename, tblStudent)

print('*** Mapping col and row values in given excel sheet***')
dlits= mapDataFields(sh, fldsStudent)

print('*** Writing data list in Json file ***')
writeToJsonFile('file.json', dlits)




