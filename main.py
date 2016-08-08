from array import *
from openpyxl import *
from collections import *
import os
import glob
import functions


def main():
   # carrega arquivo de referencia e pega os dados especificos
   wb = load_workbook(filename='C:\\Users\\Sixel1\\Desktop\\Python Excel\\Data.xlsx')
   sheet_ranges = wb.get_sheet_by_name(name='Sheet1')

   # pega valores da planilha source
   SWV_source = sheet_ranges['B6'].value
   address = sheet_ranges['B14'].value
   print("Target Address: ", address)

   # passa dados para realizar a verificação com a planilha target*/
   resultPower = check_power(address, SWV_source)

   # Matriz é sempre declarada em linhas
   ArrayList = [['1. FT - ', 'Pass'], ['2. PRI - ', 'Fail'], ['3. WDL - ', 'None'], ['4. FOTA - ', 'None'],
                ['5. POWER - ', resultPower], ['8. TD Defect Report', 'None']]

   print("Result:")

   print()
   # for para percorrer linha
   for linha in ArrayList:
      # for para percorrer coluna
      for valCol in linha:
         # printa valores
         print(valCol, end="")
      print()


main()
#
#check_dir_and_file()
