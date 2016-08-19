from array import *
from openpyxl import *
from collections import *
import os
import glob
from functions import *


def main():
   # carrega arquivo de referencia e pega os dados especificos
   wb = load_workbook(filename='C:\\Users\\alexandre.rodrigues.LGE\\Desktop\\Proposta_Inovação\\Python Excel\\Data.xlsx')
   sheet_ranges = wb.get_sheet_by_name(name='Sheet1')

   # pega valores da planilha source
   SWV_source = sheet_ranges['B6'].value
   address = sheet_ranges['B14'].value
   print("Target Address: ", address)

   # passa dados para realizar a verificação com a planilha target
   resultPower = check_power(address, SWV_source)
   result_td = check_td(address)

   # Matriz é sempre declarada em linhas
   ArrayList = [['1. FT - ', 'N/A'], ['2. PRI - ', 'N/A'], ['3. WDL - ', 'N/A'], ['4. FOTA - ', 'N/A'],
                ['5. POWER - ', resultPower], ['8. TD Defect Report - ', result_td]]

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
