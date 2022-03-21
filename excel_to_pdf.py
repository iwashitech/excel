# -*- coding: utf-8 -*-
"""
https://www.zacoding.com/post/python-excel-to-pdf/
http://oboe2uran.hatenablog.com/entry/2019/10/23/100143

https://gammasoft.jp/blog/python-parse-pdf-contents/

https://tonari-it.com/excel-vba-pdf/
https://docs.microsoft.com/ja-jp/office/vba/api/excel.pagesetup.fittopageswide

https://stackoverflow.com/questions/16683376/print-chosen-worksheets-in-excel-files-to-pdf-in-python
"""

import os
import win32com.client
import xlrd
from pywintypes import com_error

user_name = os.environ['USERPROFILE'].replace('\\', '/')

# 絶対パスで指定
WB_PATH = user_name + '/Desktop/image.xlsx'
PATH_TO_PDF = user_name + '/Desktop/image.pdf'

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

book = xlrd.open_workbook(user_name + '/Desktop/image.xlsx')
sheet_names = book.sheet_names()

try:
    print('PDFへ変換開始')

    wb = excel.Workbooks.Open(WB_PATH)

    for sheet_name in sheet_names:
        wb.WorkSheets(sheet_name).Select()
        
        # 保存
        wb.ActiveSheet.PageSetup.Zoom = False
        wb.ActiveSheet.ExportAsFixedFormat(0, user_name + '/Desktop/' + sheet_name + '.pdf')
except com_error as e:
    print('失敗しました')
else:
    print('成功しました')
finally:
    wb.Close()
    excel.Quit()