from array import *
from openpyxl import *
from collections import *
import os
import glob

# Verifica se o diretorio existe(ex 5. Power) e retorna True ou False
def check_dir(address):

   if address:
      print("\nDirectory exists: ", address)
      return True
   else:
      return False


# TODO: implementar uma função que verifica se existe uma planilha dentro de uma pasta (REVER!!)
def file_exists(address):

   new_address = address + '\\*.xlsx'

   #retorna lista de arquivos .xlsx
   result_file = glob.glob(new_address)

   if result_file:
      #retorna lista
      return result_file
   else:
      print(".xlsx not found at: ", address)
      return False

# TODO: implementar uma função que verifica abre esta planilha e valida o conteúdo dentro dela (ONGOING)
# address = endereco onde as pastas estao; SWV_source = versao de SW
def check_power(address, SWV):

   pwr_address = address + '\\5. Power'

   if check_dir(pwr_address):
      #tentativa de tratar vetor como string
      result_file = file_exists(pwr_address)
      if result_file:
         pwr_address = result_file[0]

         # abre o arquivo de acordo com o endereço passado
         wb = load_workbook(filename=pwr_address, data_only=True)
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
         if (SWV == SWV_target):
            return "Pass"
         else:
            return "Fail: SW Version does not match!"

