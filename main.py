#!/usr/bin/python
# -*- coding: utf-8 -*-

# from openpyxl import Workbook, load_workbook
import sys
import xlrd
import xlwt

sourceFileName = sys.argv[1]
groupIndex = int(sys.argv[2])

wb = xlrd.open_workbook('./%s' % sourceFileName)

sheet = wb.sheets()[0]

rows_num = sheet.nrows
cols_num = sheet.ncols

filesDict = {}

for i in range(1, rows_num):

    headers = sheet.row_values(0)
    contents = sheet.row_values(i)
    newWb = None
    sh1 = None
    contentCursor = 1

    if contents[groupIndex] in filesDict:
        contentCursor = filesDict[contents[groupIndex]][0]
        newWb = filesDict[contents[groupIndex]][1]
        sh1 = filesDict[contents[groupIndex]][2]
    else:
        newWb = xlwt.Workbook()
        sh1 = newWb.add_sheet('Sheet1')
        # header
        for j in range(cols_num):
            sh1.write(0, j, headers[j])
        filesDict[contents[groupIndex]] = [contentCursor, newWb, sh1]


    # content
    # 序号
    sh1.write(contentCursor, 0, contentCursor)
    # 序号后的列
    for j in range(1, cols_num):
        sh1.write(contentCursor, j, contents[j])

    contentCursor += 1
    filesDict[contents[groupIndex]][0] = contentCursor

for key, [contentCursor, newWb, sh] in filesDict.items():
    newWb.save('./output/%s.xls' % key)
