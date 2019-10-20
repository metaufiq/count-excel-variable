from openpyxl import *
import datetime

wb = load_workbook(r".\zona5.xlsx")


variable = {}
for sheet in wb.worksheets:
    for i in range(1,150):
        if sheet['A'+str(i)].value != None and (sheet['A'+str(i+1)].value != None or sheet['B'+str(i+1)].value != None):
            data = str(sheet['A'+str(i)].value).split(' ')
            if data[0] =='KODE' and (data[1] == 'TARIF' or  data[1] == 'TARIP'):
                
                if data[2].upper() in variable:
                    variable[data[2].upper()] += 1
                else:
                    variable[data[2].upper()] = 1

print(variable)