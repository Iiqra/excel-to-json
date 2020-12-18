import xlrd
from collections import OrderedDict
import json

dbFilename = 'SententialDB.xlsx'
indexFormula = 0
indexLanguage = 1
indexTense = 2
indexSentence = 3
indexTranslation = 4


fldsFormula = 'ID, TenseID, LanguageID, Formula'
fldsLanguage = 'ID, Language, PossibleMappings'
fldsTense = 'ID, Type'
fldsSentence = 'ID, Sentence'
fldsTranslation = 'ID, TenseID, LanguageID, SentenceID, Translation'

flFormula = "Formula.json"
flLanguage = "Language.json"
flTense = "Tense.json"
flSentence = "Sentence.json"
flTranslation = "Translation.json"


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

def excelSheetToJsonDataConversion(dbFilename, tableIndex, tableFields, jsonFilename):
    print('*** Reading Data from Excel Sheet***')
    sh = readExcelSheet(dbFilename, tableIndex)

    print('*** Mapping col and row values in given excel sheet***')
    dlits= mapDataFields(sh, tableFields)

    print('*** Writing data list in Json file ***')
    writeToJsonFile(jsonFilename, dlits)


#1. For table Formula
excelSheetToJsonDataConversion(dbFilename, indexFormula,fldsFormula, flFormula)

#2. For table Language
excelSheetToJsonDataConversion(dbFilename, indexLanguage,fldsLanguage, flLanguage)

#3. For table Tense
excelSheetToJsonDataConversion(dbFilename,indexTense ,fldsTense, flTense)

#4. For table Sentence
excelSheetToJsonDataConversion(dbFilename, indexSentence,fldsSentence, flSentence)

#5. For table Translation
excelSheetToJsonDataConversion(dbFilename, indexTranslation,fldsTranslation, flTranslation)









