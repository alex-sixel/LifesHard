from array import *
from openpyxl import *
from collections import *
import os
import glob


# Verifica se existem diretorios e planilhas
# TODO: implementar uma funçao que verifica se a pasta existe (ex 5. Power) e returna true or false
# TODO: implementar uma função que verifica se existe uma planilha dentro de uma pasta
# TODO]: implementar uma função que verifica abre esta planilha e valida o conteúdo dentro dela (ONGOING)

def check_dir_and_file(address, folder_to_check ):
   resultDir = os.path.isdir(address,'\\',folder_to_check)
   # if true (diretorio existe)
   if resultDir:
      print("\nDirectory exists: ", resultDir)

      resultFile = glob.glob(address,'\\',folder_to_check,'\\*.xlsx')

      # if true (arquivos existem)
      if resultFile:
         print("Extensions *.xlsx exists: ", resultFile)
         return True
      else:
         print("Extension not found!")
         return False
   else:
      print("Directory not found!")
      return False


def check_power(address, SWV_source):
   # abre o arquivo de acordo com o endereço passado
   wb = load_workbook(filename=address, data_only=True)
   sheet_ranges = wb.get_sheet_by_name(name='Checklist Report')

   # percorre na coluna K e mostra os valores
   colK = sheet_ranges.columns[10]  # 0 = A, 1 = B...10 = K
   for valueK in colK:
      #verifica se valor da celula é NG/Fail
      if (valueK.value == "NG" or valueK.value == "ng" or valueK.value == "Fail" or valueK.value == "fail"):
         return "Fail result found at Test Satus column!"

   # Converte valor da celula E8 para INT e verifica se ha algum fail nesta celula (valor diferente de 0)
   E8Int = int(sheet_ranges['E8'].value)
   if (E8Int != 0):
      return "Fail: " + str(sheet_ranges['E8'].value) + " error(s) found at result table"

   #Verifica versao de SW
   SWV_target = sheet_ranges['B5'].value
   if (SWV_source == SWV_target):
      # print("PASS")
      return "Pass"
   else:
      return "Fail: SW Version do not match!"


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
