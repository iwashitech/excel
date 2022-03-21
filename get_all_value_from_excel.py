# -*- coding: utf-8 -*-
"""

https://stackoverflow.com/questions/8909342/python-xlrd-how-to-convert-an-extracted-value
https://note.nkmk.me/python-xlrd-xlwt-usage/

"""

import os
import xlrd
import pandas as pd

user_name = os.environ['USERPROFILE'].replace('\\', '/')

book = xlrd.open_workbook(user_name + '/Desktop/value.xlsx')
sheet_names = book.sheet_names()

for sheet_name in sheet_names:
    sheet = book.sheet_by_name(sheet_name)
    
    cells = []
    for i in range(sheet.nrows):
        cell_list = sheet.row_values(rowx=i,start_colx=0,end_colx=None)
        
        for cell in cell_list:
            if cell != '':
                cells.append(cell)
    df = pd.DataFrame(cells)
    df.to_csv(user_name + '/Desktop/' + sheet_name + '.csv', header=False, index=False)