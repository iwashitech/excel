# -*- coding: utf-8 -*-
"""

https://github.com/glexey/excel2img
https://stackoverflow.com/questions/26521266/using-pandas-to-pd-read-excel-for-multiple-worksheets-of-the-same-workbook#answer-46081870

"""

import os
import pandas as pd
import excel2img

user_name = os.environ['USERPROFILE'].replace('\\', '/')

xls = pd.ExcelFile(user_name + '/Desktop/image.xlsx')

for sheet_name in xls.sheet_names:
    excel2img.export_img(user_name + '/Desktop/image.xlsx', user_name + '/Desktop/' + sheet_name + '.png', sheet_name, None)
    # IF Exception Occur => Failed locating used cells or single object to print on page
    #excel2img.export_img(user_name + '/Desktop/image.xlsx', user_name + '/Desktop/' + sheet_name + '.png', '', sheet_name + '!A1:AD38')