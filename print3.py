# -*- coding: utf-8 -*-
import openpyxl

file = 'MarketingAnalystNames.xlsx'
file2 = 'MarketingAnalystNames22.xlsx'
wb = openpyxl.load_workbook(filename=file)

# Seleciono la Hoja
ws = wb.get_sheet_by_name('Sheet1')

# Valores a Insertar
ws['A3'] = 42
ws['A4'] = 142

print(len(wb.sheetnames))
sheets = wb.sheetnames

for sh in range (len(wb.sheetnames)):
        sheet = wb.get_sheet_by_name(sheets[sh])
        for i in range(1,sheet.max_row):
                for j in range(1,sheet.max_column):
                                print(sh,i,j,sheet.cell(i,j).value)
                                sheet.cell(i,j).value = 0
                                

# Escribirmos en el Fichero
wb.save(file2)
