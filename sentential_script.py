import xlrd
from collections import OrderedDict
import json

dbFilename = 'SententialDB.xlsx'
indexFormula = 0
indexLanguage = 1
indexTense = 2
indexSentence = 3
indexTranslation = 4


fldsFormula = 'ID,TenseID,LanguageID,Formula'
fldsLanguage = 'ID,Language'
fldsTense = 'ID,Type'
fldsSentence = 'ID,Sentence'
fldsTranslation = 'ID,TenseID,LanguageID,SentenceID,Translation'

flFormula = "formula.json"
flLanguage = "language.json"
flTense = "tense.json"
flSentence = "sentence.json"
flTranslation = "translation.json"

def read_excel_sheet(filename, index):
     wb = xlrd.open_workbook(filename)
     sheet = wb.sheet_by_index(index)
     return sheet

def map_data_fields(sheet, fields):
     data_list = []
     field_list = fields.split(',')
     for rownum in range(1, sheet.nrows):
          data = OrderedDict()
          i = 0
          row_values = sheet.row_values(rownum)
          while(i < len(field_list)):
               data[field_list[i]] = row_values[i]
               i +=1
          data_list.append(data)

     return data_list

def write_to_json_file(file_name, data_list):
     with open(file_name, "w", encoding="utf-8") as write_json_file:
           json.dump(data_list, write_json_file, indent=4, default=str)

def excel_sheet_to_json_data_conversion(db_file_name, table_index, table_fields, json_file_name):
    print('*** Reading Data from Excel Sheet***')
    sh = read_excel_sheet(db_file_name, table_index)

    print('*** Mapping col and row values in given excel sheet***')
    dlits = map_data_fields(sh, table_fields)

    print('*** Writing data list in Json file ***')
    write_to_json_file(json_file_name, dlits)

#1. For table Formula
excel_sheet_to_json_data_conversion(dbFilename, indexFormula,fldsFormula, flFormula)

#2. For table Language
excel_sheet_to_json_data_conversion(dbFilename, indexLanguage,fldsLanguage, flLanguage)

#3. For table Tense
excel_sheet_to_json_data_conversion(dbFilename,indexTense ,fldsTense, flTense)

#4. For table Sentence
excel_sheet_to_json_data_conversion(dbFilename, indexSentence,fldsSentence, flSentence)

#5. For table Translation
excel_sheet_to_json_data_conversion(dbFilename, indexTranslation,fldsTranslation, flTranslation)
