#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: Jan Yang
@software: PyCharm Community Edition
@time: 2018/1/19 11:45
"""


import xlsxwriter

workbook = xlsxwriter.Workbook('chart_secondary_axis.xlsx')
worksheet = workbook.add_worksheet()

data = [
    [2, 3, 4, 5, 6, 7],
    [10, 40, 50, 20, 10, 50],
]

worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])

chart = workbook.add_chart({'type': 'line'})

# Configure a series with a secondary axis.
chart.add_series({
    'values': '=Sheet1!$A$2:$A$7',
})

# Configure a primary (default) Axis.
chart.add_series({
    'values': '=Sheet1!$B$2:$B$7',
})

chart.add_series({
    'values': '=Sheet1!$A$2:$A$7',
    'y2_axis': True,
})


chart.set_legend({'position': 'none'})

chart.set_y_axis({'name': 'Primary Y axis'})
chart.set_y2_axis({'name': 'Secondary Y axis'})

worksheet.insert_chart('D2', chart)

workbook.close()