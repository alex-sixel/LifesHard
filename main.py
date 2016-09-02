from array import *
from openpyxl import *
from collections import *
import os
import glob
import time
from functions import *


def main():
   # carrega arquivo de referencia e pega os dados especificos
   wb = load_workbook(filename='C:\\Users\\alexandre.rodrigues.LGE\\Desktop\\Proposta_Inovação\\Python Excel\\Data.xlsx')
   sheet_ranges = wb.get_sheet_by_name(name='Sheet1')

   # pega valores da planilha source
   model = sheet_ranges['B2'].value
   suffix = sheet_ranges['B3'].value
   sw_phase = sheet_ranges['B5'].value
   SWV_source = sheet_ranges['B6'].value
   address = sheet_ranges['B14'].value
   print("Target Address: ", address)

   start_time = time.time()

   # passa dados para realizar a verificação com a planilha target
   resultPower = check_power(address, SWV_source)
   result_td = check_td(address)
   result_pri = check_pri(address, SWV_source, model, suffix)
   # Matriz é sempre declarada em linhas
   ArrayList = [['1. FT - ', 'N/A'], ['2. PRI - ', result_pri], ['3. WDL - ', 'N/A'], ['4. FOTA - ', 'N/A'],
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

   print("\n\nFinished in %0.2f seconds" %(time.time() - start_time))
main()
#
#check_dir_and_file()
