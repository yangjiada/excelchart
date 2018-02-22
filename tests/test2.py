#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: Jan Yang
@software: PyCharm Community Edition
@time: 2018/2/14 15:04
"""


import xlsxwriter

workbook = xlsxwriter.Workbook('chart_test.xlsx')
worksheet = workbook.add_worksheet()

# Create a new Chart object.
chart = workbook.add_chart({'type': 'line'})

# Write some data to add to plot on the chart.
data = [
    [1, 2, 3, 4, 5],
    [2, 4, 6, 8, 10],
    [3, 6, 9, 12, 15],
]

worksheet.write_column('A1', data[0])
worksheet.write_column('B1', data[1])
worksheet.write_column('C1', data[2])

# Configure the chart. In simplest case we add one or more data series.
chart.add_series({'values': '=Sheet1!$A$1:$A$5'})
chart.add_series({'values': '=Sheet1!$B$1:$B$5'})
chart.add_series({'values': '=Sheet1!$C$1:$C$5'})
# chart.set_y_axis({'min': 2, 'max': 7})
# chart.set_x_axis({
#     'name': '标题ABC123'
# })
# chart.set_x_axis({
#     'num_font': {'name': '华文彩云'}
# })


chart.set_x_axis({'interval_tick': 2})
chart.set_x_axis({'interval_unit': 2})
# Insert the chart into the worksheet.
worksheet.insert_chart('A7', chart)

workbook.close()