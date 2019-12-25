# Program extracting first column 
import xlrd
import xlwt
from xlutils.copy import copy

import pandas as pd

sales_rep_names = pd.read_excel("SalesRepNames.xlsx",sheet_name=None)


loc = ("MarketingAnalystNames.xlsx")
master = ("MarketingAnalystNames_Merge.xlsx") 

wb = xlrd.open_workbook(loc, on_demand=True)
ms = xlrd.open_workbook(master, on_demand=True)


n_ms = copy(ms)

for sh in range (len(wb.sheet_names())):
        sheet = wb.sheet_by_index(sh)
        ms_sheet = ms.sheet_by_index(sh)
        n_ms_sheet = n_ms.get_sheet(sh) 
        for i in range(1,sheet.nrows):
                for j in range(sheet.ncols):
                        if (ms_sheet.cell_value(i,j) == "") & (sheet.cell_value(i,j) != "") :
                                print(sh,i,j,sheet.cell_value(i,j))
                                n_ms_sheet.write(i,j,sheet.cell_value(i,j))
                                
n_ms.save("MarketingAnalystNames_Merge2.xls")
