#!/usr/bin/python
# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf8')

import xlrd

data = xlrd.open_workbook('Input.xlsx')
table = data.sheet_by_index(0)
cell = table.cell(0,0)

import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Output.xlsx')
worksheet = workbook.add_worksheet()

format = workbook.add_format()
format.set_align('vcenter')

row = 0
col = 0
row_idx = 0
col_idx = 0
num_cols = table.ncols   # Number of columns
for row_idx in range(0, table.nrows):    # Iterate through rows
    if row % 7 == 0:
        row += 1
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = table.cell(row_idx, col_idx)  # Get cell object by row, col
        if table.cell(row_idx,2).value != 0:
            worksheet.write(row, col_idx, cell_obj.value, format)
    if table.cell(row_idx,2).value != 0:
        row += 1


workbook.close()
