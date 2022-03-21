# -*- coding: utf-8 -*-
"""

"""

import xlwings as xw
import os

user_name = os.environ['USERPROFILE'].replace('\\', '/')

#Excelファイルを用意する必要がある
wb = xw.Book(user_name + '/Desktop/output.xlsx')

app = xw.apps.active

#記入している必要がある
cell_list = xw.Range('A1:A100').value

for i in range(10):
    xw.Range('B' + str(i+1)).value = cell_list[i] * 10

wb.save(user_name + '/Desktop/output.xlsx')
app.quit()