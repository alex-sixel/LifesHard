from array import *
from openpyxl import *
from openpyxl.cell import *
from collections import *
import os
import glob

# Verifica se o diretorio existe(ex 5. Power) e retorna True ou False
def check_dir(address):

   if address:
     #print("\nDirectory exists: ", address)
      return True
   else:
      return False


# Retorna lista de arquivos .xlsx dentro do endereço recebido
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

# TODO: IMPLEMENTAR
def check_pri():
   return True

# TODO: IMPLEMENTAR COMPARAÇÃO COM PREVIOUS
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

# TODO: IMPLEMENTAR
# Procurar pela coluna 'Status', verificar se todas as células estão fechadas; Procurar pela coluna 'Model' e ver se é o mesmo modelo (OK)
# address = endereco onde as pastas estao;
def check_td(address):

   td_folder = address + '\\8. TD Defect Report'

   if check_dir(td_folder):
      file_list = file_exists(td_folder)
      if file_list:
         # file_list[0] = primeiro arquivo encontrado
         td_file = file_list[0]

         # abre o arquivo de acordo com o endereço passado
         wb = load_workbook(filename=td_file, data_only=True)
         current_sheet = wb.get_sheet_by_name(name='Sheet1')

         col_range = current_sheet['A1:AJ1']
         #percore a linha1
         for row in col_range:
            # Percorre colunaX, até achar o valor "Status"
            for cell in row:
               #print("At ", cell.column, cell.row, " cell value is: ", cell.value)
               if cell.value == "Status":
                  print("Cell value: ", cell.value)

                  col_index = column_index_from_string(cell.column)

                  for cell_status in current_sheet.columns[col_index-1]:
                     # TODO: ** INSERIR ANY AQUI **
                     if(cell_status.value == "Status" or cell_status.value == "Closed" or cell_status.value == "Closed.Not a bug" or cell_status.value == "Closed.Withdrawn" or cell_status.value == "Closed.Deferred"):
                        print("Within Col Status: ", cell_status.value)
                     else:
                        print("Error! not all issues are properly closed!")
                        return "Fail"
                  print("All good!")
                  return "Pass"
