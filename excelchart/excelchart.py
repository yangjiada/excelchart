#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: Jan Yang
@software: PyCharm Community Edition
"""


import pandas as pd
import string
import xlsxwriter


class Chart(object):
    """ The chart of class.

    """
    def __init__(self, workbook, frame, sheet_name=None, chart_type=None, subtype=None, date_parser=False):
        """

        :param workbook: string
        :param frame: DataFrame
        :param sheet_name: string, default None
        :param chart_type: string, default None
        :param subtype: string, default None
        :param date_parser: bool, default False
        :return:
        """
        self.workbook = workbook
        self.frame = frame
        self.sheet_name = sheet_name
        self.chart_type = chart_type
        self.subtype = subtype

        self.uppercase = string.ascii_uppercase
        self.data = frame.values
        self.shape = frame.shape

        self.bold = workbook.add_format({'bold': 1})
        self.date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

        # Create a workbook and write the data.
        self.worksheet = self.workbook.add_worksheet(sheet_name)

        # x_
        self.x_axis_params = {}
        self.y_axis_params = {}

        if sheet_name is None:
            self.sheet_name = self.worksheet.name

        # Write data to worksheet.
        self.worksheet.write_row('A1', self.frame.columns, self.bold)

        for col in range(self.shape[1]):
            self.worksheet.write_column(1, col, self.data[:, col],
                                        self.date_format if date_parser and col == 0 else None)

        # Create a Chart object.
        self.chart = self.workbook.add_chart({'type': chart_type, 'subtype': subtype})

    def add_series(self, data_labels, overlap=0, gap=150):
        """ Add one or more data series.

        :param data_labels: bool
        :param overlap: int, default 0
        :param gap: int, default 150
        :return:
        """

        for num in range(1, self.shape[1]):
            '[sheet_name, first_row, first_col, last_row, last_col]'
            self.chart.add_series({
                'name': [self.sheet_name, 0, num, 0, num],
                'categories': [self.sheet_name, 1, 0, self.shape[0], 0],
                'values': [self.sheet_name, 1, num, self.shape[0], num],
                'data_labels': {'value': data_labels},
                'overlap': overlap,
                'gap': gap
            })

    def set_size(self, width=480, height=350, x_scale=1, y_scale=1, x_offset=0, y_offset=0):
        """ Set the dimensions of the chart.

        :param width: int, default 480
        :param height: int, default 350
        :param x_scale: int, default 5
        :param y_scale: int, default 3
        :param x_offset: int, default 0
        :param y_offset: int, default 0
        :return:
        """

        self.chart.set_size({
            'width': width,
            'height': height,
            'x_scale': x_scale,
            'y_scale': y_scale,
            'x_offset': x_offset,
            'y_offset': y_offset
        })

    def set_title(self, title, font='Arial', size=16, bold=True, italic=False, underline=False,
                  color='black', rotation=0, overlay=False, layout=None):
        """ Set the chart title.

        :param title: string
        :param font: string
        :param size: int, default 16
        :param bold: bool, default True
        :param italic: bool, default False
        :param underline: bool, default False
        :param color: string, default
        :param rotation: int, default 0
        :param overlay: bool, default False
        :param layout: tuple, default None
            layout -> (x, y): x and y as a percentage and 0 < x <= 1
        :return:
        """

        if title:
            self.chart.set_title({
                'name': title,
                'name_font': {
                    'name': font,
                    'size': size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                },
                'overlay': overlay,
                'layout': {'x': layout[0], 'y': layout[1]} if layout else None
            })
        # else:0

        #     self.chart.set_title({'none': not title})

    def set_legend(self, legend, font='Arial', font_size=10, bold=False, italic=False, underline=False,
                   rotation=0, color='black', delete_series=None, layout=None):
        """ Set the chart legend.

        :param legend: string, top bottom left right overlay_left overlay_right none
        :param font: string
        :param font_size: int, default 16
        :param bold: bool, default False
        :param italic: bool, default False
        :param underline: bool, default False
        :param rotation: int, default 0
        :param color: string, default 'black'
        :param delete_series: list, default None
        :param layout: tuple, default None
            layout -> (x, y, width, height)
        :return:
        """
        if legend:
            self.chart.set_legend({
                'position': legend,
                'font': {
                    'name': font,
                    'size': font_size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                },
                'delete_series': delete_series,
                'layout': {'x': layout[0], 'y': layout[1], 'width': layout[2], 'height': layout[3]} if layout else None
            })

    def set_chart_area(self, border=False, border_color='black',  width=0.75, border_transparency=50, dash_type='solid',
                       fill=False, fill_color='white', fill_transparency=0, pattern=None, gradient=None):
        """ Set the chart area.

        :param border: bool, default False
        :param border_color: string, default 'black'
        :param width: float or int, default 0.75
        :param border_transparency: int, default 50
        :param dash_type: string, default 'solid'
        :param fill: bool, default False
        :param fill_color: string, default 'white'
        :param fill_transparency: int, default 50
        :param pattern:
        :param gradient:
        :return:
        """
        _ = pattern, gradient

        self.chart.set_chartarea({
            'border': {
                'color': border_color,
                'width': width,
                'transparency': border_transparency,
                'dash_type': dash_type
            } if border else {'none': True},
            'fill': {
                'color': fill_color,
                'transparency': fill_transparency,
            } if fill else {'none': True}
        })

    def set_plot_area(self, border=False, border_color='black', width=0.75, border_transparency=50, dash_type='solid',
                      fill=False, fill_color='white', fill_transparency=0, pattern=None, gradient=None, layout=None):
        """ Set the plot area.

        :param border: bool, default False
        :param border_color: string, default 'black'
        :param width: float or int, default 0.75
        :param border_transparency: int, default 50
        :param dash_type: string, default 'solid'
        :param fill: bool, default False
        :param fill_color: string, default 'white'
        :param fill_transparency: int, default 0
        :param pattern:
        :param gradient:
        :param layout: tuple, default None
        :return:
        """

        _ = pattern, gradient

        self.chart.set_plotarea({
            'border': {
                'color': border_color,
                'width': width,
                'transparency': border_transparency,
                'dash_type': dash_type
            } if border else {'none': True},
            'fill': {
                'color': fill_color,
                'transparency': fill_transparency
            } if fill else {'none': True},
            'layout': {'x': layout[0], 'y': layout[1], 'width': layout[2], 'height': layout[3]} if layout else None
        })

    def set_style(self, style_id):
        """ Set the chart style type.

        :param style_id: int
        :return:
        """
        if style_id:
            self.chart.set_style(style_id)

    def set_table(self, horizontal=True, vertical=True, outline=True, show_keys=False, font='Arial',
                  font_size=10, bold=False, italic=False, underline=False, rotation=0, color='black'):
        """ Set data table.

        :param horizontal: bool, default True
        :param vertical: bool, default True
        :param outline: bool, default True
        :param show_keys: bool, default False
        :param font: string, default 'Arial'
        :param font_size: int, default 10
        :param bold: bool, default False
        :param italic: bool, default False
        :param underline: bool, default False
        :param rotation: int, default 0
        :param color: string, default 'black'
        :return:
        """

        self.chart.set_table({
            'horizontal': horizontal,
            'vertical': vertical,
            'outline': outline,
            'show_keys': show_keys,
            'font': {
                'name': font,
                'size': font_size,
                'bold': bold,
                'italic': italic,
                'underline': underline or None,
                'rotation': rotation,
                'color': color,
            }
        })

    def set_up_down_bars(self, up_border_color='black', up_width=0.75, up_border_transparency=50, up_dash_type='solid',
                         up_fill_color='green', up_fill_transparency=50, down_border_color='black', down_width=0.75,
                         down_border_transparency=50, down_dash_type='solid', down_fill_color='red',
                         down_fill_transparency=0):
        """ Set properties for the chart up-down bars.

        :param up_border_color: string, default 'black'
        :param up_width: float or int, default 0.75
        :param up_border_transparency: int, default 50
        :param up_dash_type: string, default 'solid'
        :param up_fill_color: string, default 'green'
        :param up_fill_transparency: int, default 50
        :param down_border_color: string, default 'black'
        :param down_width: float or int, default 0.75
        :param down_border_transparency: int, default 50
        :param down_dash_type: string, default 'solid'
        :param down_fill_color: string, default 'red'
        :param down_fill_transparency: int, default 0
        :return:
        """

        self.chart.set_up_down_bars({
            'up': {
                'border': {
                    'color': up_border_color,
                    'width': up_width,
                    'transparency': up_border_transparency,
                    'dash_type': up_dash_type
                },
                'fill': {
                    'color': up_fill_color,
                    'transparency': up_fill_transparency,
                }
            },
            'down': {
                'border': {
                    'color': down_border_color,
                    'width': down_width,
                    'transparency': down_border_transparency,
                    'dash_type': down_dash_type
                },
                'fill': {
                    'color': down_fill_color,
                    'transparency': down_fill_transparency,
                }
            }
        })

    def set_drop_lines(self, color='black', width=0.75, dash_type='solid', transparency=50):
        """ Set properties for the chart drop lines.

        :param color: string, default 'black'
        :param width: float or int, default 0.75
        :param dash_type: string, default 'solid'
            solid round_dot square_dot dash dash_dot long_dash long_dash_dot long_dash_dot_dot
        :param transparency: int, default 50
        :return:
        """

        #
        self.chart.set_drop_lines({
            'line': {
                'color': color,
                'width': width,
                'dash_type': dash_type,
                'transparency': transparency
            }
        })

    def set_high_low_lines(self, color='black', width=0.75, dash_type='solid', transparency=50):
        """ Set properties for the chart high-low lines.

        :param color: string, default 'black'
        :param width: float or int, default 0.75
        :param dash_type: string, default 'solid'
        :param transparency: int, default 50
        :return:
        """
        self.chart.set_high_low_lines({
            'line': {
                'color': color,
                'width': width,
                'dash_type': dash_type,
                'transparency': transparency
            }
        })

    def show_blanks_as(self, show='gap'):
        """ Set displaying blank data in the chart.

        :param show: string, gap zero span default 'gap'
        :return:
        """
        self.chart.show_blanks_as(show)

    def show_hidden_data(self):
        """ Display data on charts from hidden rows or columns.

        :return:
        """
        self.show_hidden_data()

    def set_rotation(self, rotation=0):
        """ Set the Pie/Doughnut chart rotation.

        :param rotation: int, default 0
            0 <= rotation <= 360
        :return:
        """
        self.chart.set_rotation(rotation)

    def set_hole_size(self, size):
        """ Set the Doughnut chart hole size.

        :param size: int
            10 <= size <= 90
        :return:
        """

        self.set_hole_size(size)

    def set_x_axis(self, name=None, name_font='Arial', name_size=12, name_bold=True, name_italic=False,
                   name_underline=False, name_rotation=0, name_color='black', name_layout=None, label_font='Arial',
                   label_size=10, label_bold=False, label_italic=False, label_underline=False, label_rotation=0,
                   label_color='black', num_format=None, line=True, line_color='black', width=0.75, dash_type='solid',
                   line_transparency=50, fill=False, fill_color='white', fill_transparency=0):
        self.chart.set_x_axis({
            'name': name,
            'name_font': {
                    'name': name_font,
                    'size': name_size,
                    'bold': name_bold,
                    'italic': name_italic,
                    'underline': name_underline or None,
                    'rotation': name_rotation,
                    'color': name_color,
                },
            'name_layout': {'x': name_layout[0], 'y': name_layout[1]} if name_layout else None,
            'num_font': {
                'name': label_font,
                'size': label_size,
                'bold': label_bold,
                'italic': label_italic,
                'underline': label_underline or None,
                'rotation': label_rotation,
                'color': label_color,
            },
            'num_format': num_format,
            'line': {
                'color': line_color,
                'width': width,
                'dash_type': dash_type,
                'transparency': line_transparency
            } if line else {'none': True},
            'fill': {
                'color': fill_color,
                'transparency': fill_transparency,
            } if fill else {'none': True}
        })

    def set_x_tick_mark(self, major_type=None, minor_type=None):
        """ Set x axis tick mark type.

        :param major_type: string, inside outside cross, default 'none'
        :param minor_type: string, inside outside cross, default 'none'
        :return:
        """
        self.chart.set_x_axis({
            'major_tick_mark': major_type or 'none',
            'minor_tick_mark': minor_type or 'none'
        })

    def set_y_tick_mark(self, major_type=None, minor_type=None):
        """ Set y axis tick mark type.

        :param major_type: string, inside outside cross, default 'none'
        :param minor_type: string, inside outside cross, default 'none'
        :return:
        """
        self.chart.set_y_axis({
            'major_tick_mark': major_type or 'none',
            'minor_tick_mark': minor_type or 'none'
        })

    def set_x_title(self, title, font='Arial', size=10, bold=True, italic=False, underline=False, rotation=0,
                    color='black', layout=None):
        """ Set the chart x axis title.

        :param title: string
        :param font: string, default 'Arial'
        :param size: int, default 10
        :param bold: bool, default True
        :param italic: bool, default False
        :param underline: bool, default False
        :param rotation: int, default 0
        :param color: string, default 'black'
        :param layout: tuple, default None
        :return:
        """

        if title:
            self.x_axis_params.update({
                'name': title,
                'name_font': {
                    'name': font,
                    'size': size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                },
                'name_layout': {'x': layout[0], 'y': layout[1]} if layout else None,
            })

    def set_y_title(self, title, font='Arial', size=10, bold=True, italic=False, underline=False, rotation=0,
                    color='black', layout=None):
        """ Set the chart y axis title.

        :param title: string
        :param font: string, default 'Arial'
        :param size: int, default 10
        :param bold: bool, default True
        :param italic: bool, default False
        :param underline: bool, default False
        :param rotation: int, default 0
        :param color: string, default 'black'
        :param layout: tuple, default None
        :return:
        """
        if title:
            self.y_axis_params.update({
                'name': title,
                'name_font': {
                    'name': font,
                    'size': size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                },
                'name_layout': {'x': layout[0], 'y': layout[1]} if layout else None
            })

    def set_x_label(self, font='Arial', size=10, bold=False, italic=False, underline=False, rotation=0,
                    color='black', num_format=None, interval_unit=None, position='next_to'):
        """ Set the chart x axis label params.

        :param font: string, default 'Arial'
        :param size: int, default 10
        :param bold: bool, default False
        :param italic: bool, default False
        :param underline: bool, default False
        :param rotation: int, default 0
        :param color: string, default 'black'
        :param num_format: string, default None
        :param interval_unit: int default None
        :param position: string, next_to high low none, default 'next_to'
        :return:
        """
        self.x_axis_params.update({
            'num_font': {
                'name': font,
                'size': size,
                'bold': bold,
                'italic': italic,
                'underline': underline or None,
                'rotation': rotation,
                'color': color,
            },
            'num_format': num_format,
            'interval_unit': interval_unit,
            'label_position': position
        })

    def set_y_label(self, font='Arial', size=10, bold=False, italic=False, underline=False, rotation=0,
                    color='black', num_format=None, interval_unit=None, position='next_to'):
        """ Set the chart y axis label params.

        :param font: string, default 'Arial'
        :param size: int, default 10
        :param bold: bool, default False
        :param italic: bool, default False
        :param underline: bool, default False
        :param rotation: int, default 0
        :param color: string, default 'black'
        :param num_format: string, default None
        :param interval_unit: int default None
        :param position: string, next_to high low none, default 'next_to'
        :return:
        """
        self.y_axis_params.update({
            'num_font': {
                'name': font,
                'size': size,
                'bold': bold,
                'italic': italic,
                'underline': underline or None,
                'rotation': rotation,
                'color': color,
            },
            'num_format': num_format,
            'interval_unit': interval_unit,
            'label_position': position
        })

    def set_x_tick(self, interval_unit=None, major_type='outside', minor_type='none'):
        """ Set the axis x tick mark type.

        :param interval_unit: int, default None
        :param major_type: string, default 'outside'
        :param minor_type: string, default 'none'
        :return:
        """
        self.x_axis_params.update({
            'interval_tick': interval_unit,
            'major_tick_mark': major_type,
            'minor_tick_mark': minor_type
        })

    def set_y_tick(self, interval_unit=None, major_type='outside', minor_type='none'):
        """ Set the axis y tick mark type.

        :param interval_unit: int, default None
        :param major_type: string, default 'outside'
        :param minor_type: string, default 'none'
        :return:
        """
        self.y_axis_params.update({
            'interval_tick': interval_unit,
            'major_tick_mark': major_type,
            'minor_tick_mark': minor_type
        })

    def set_x_limit(self, limit):
        """ Set the minimum and maximum for the x axis.

        :param limit: tuple
            limit -> (min, max)
        :return:
        """
        if limit:
            self.x_axis_params.update({'min': limit[0], 'max': limit[1]})

    def set_y_limit(self, limit):
        """ Set the minimum and maximum for the y axis.

        :param limit: tuple
            limit -> (min, max)
        :return:
        """
        if limit:
            self.y_axis_params.update({'min': limit[0], 'max': limit[1]})

    def set_x_unit(self, major_unit, minor_unit):
        """ Set the x axis major unit and minor unit.

        :param major_unit: int
        :param minor_unit: int
        :return:
        """
        self.x_axis_params.update({
            'major_unit': major_unit,
            'minor_unit': minor_unit
        })

    def set_y_unit(self, major_unit, minor_unit):
        """ Set the y axis major unit and minor unit.

        :param major_unit: int
        :param minor_unit: int
        :return:
        """
        self.chart.set_y_axis({
            'major_unit': major_unit,
            'minor_unit': minor_unit
        })

    def set_x_grid(self, major=False, minor=False, major_color='black', major_width=0.75, major_dash_type='solid',
                   major_transparency=50, minor_color='black', minor_width=0.75, minor_dash_type='solid',
                   minor_transparency=50):
        """ Set the x axis grid lines.

        :param major: bool, default False
        :param minor: bool, default False
        :param major_color: string, default 'black'
        :param major_width: float or int, default 0.75
        :param major_dash_type: string, default 'solid'
        :param major_transparency: int, default 50
        :param minor_color: string, default 'black'
        :param minor_width: float or int, default 0.75
        :param minor_dash_type: string, default 'solid'
        :param minor_transparency: int, default 50
        :return:
        """
        self.x_axis_params.update({
            'major_gridlines': {
                'visible': major,
                'line': {
                    'color': major_color,
                    'width': major_width,
                    'dash_type': major_dash_type,
                    'transparency': major_transparency
                }
            },
            'minor_gridlines': {
                'visible': minor,
                'line': {
                    'color': minor_color,
                    'width': minor_width,
                    'dash_type': minor_dash_type,
                    'transparency': minor_transparency
                }
            }
        })

    def set_y_grid(self, major=False, minor=False, major_color='black', major_width=0.75, major_dash_type='solid',
                   major_transparency=50, minor_color='black', minor_width=0.75, minor_dash_type='solid',
                   minor_transparency=50):
        """ Set the y axis grid lines.

        :param major: bool, default False
        :param minor: bool, default False
        :param major_color: string, default 'black'
        :param major_width: float or int, default 0.75
        :param major_dash_type: string, default 'solid'
        :param major_transparency: int, default 50
        :param minor_color: string, default 'black'
        :param minor_width: float or int, default 0.75
        :param minor_dash_type: string, default 'solid'
        :param minor_transparency: int, default 50
        :return:
        """
        self.y_axis_params.update({
            'major_gridlines': {
                'visible': major,
                'line': {
                    'color': major_color,
                    'width': major_width,
                    'dash_type': major_dash_type,
                    'transparency': major_transparency
                }
            },
            'minor_gridlines': {
                'visible': minor,
                'line': {
                    'color': minor_color,
                    'width': minor_width,
                    'dash_type': minor_dash_type,
                    'transparency': minor_transparency
                }
            }
        })

    def set_x_reverse(self):
        """ Reverse the order of the x axis categories or values.

        :return:
        """
        self.x_axis_params.update({'reverse': True})

    def set_y_reverse(self):
        """ Reverse the order of the y axis categories or values.

        :return:
        """
        self.y_axis_params.update({'reverse': True})

    def set_x_crossing(self, category):
        """ Set the position where the y axis will cross the x axis.

        :param category: int or string, max
        :return:
        """
        self.x_axis_params.update({'crossing': category})

    def set_y_crossing(self, category):
        """ Set the position where the x axis will cross the y axis.

        :param category: int or string, max
        :return:
        """
        self.y_axis_params.update({'crossing': category})

    def set_x_position(self, category):
        """ Set the x axis position.

        :param category: string, on_tick or between, default 'between'
        :return:
        """
        self.x_axis_params.update({'position_axis': category})

    def set_y_position(self, category):
        """ Set the y axis position.

        :param category: string, on_tick or between, default 'between'
        :return:
        """
        self.y_axis_params.update({'position_axis': category})

    def set_x_log(self, base=10):
        """ Set the x axis logarithmic scale.

        :param base: int, default 10
        :return:
        """
        self.chart.set_x_axis({'log_base': base})

    def set_y_log(self, base=10):
        """ Set the y axis logarithmic scale.

        :param base: int, default 10
        :return:
        """
        self.chart.set_y_axis({'log_base': base})

    def set_x_visible(self, visible=True):
        """ Configure the visibility of the x axis.

        :param visible:
        :return:
        """
        self.x_axis_params.update({'visible': visible})

    def set_y_visible(self, visible=True):
        """ Configure the visibility of the y axis.

        :param visible:
        :return:
        """
        self.y_axis_params.update({'visible': visible})

    def set_x_text_axis(self):
        """ Set x axis as a text axis.

        :return:
        """
        self.x_axis_params.update({'text_axis': True})

    def set_y_text_axis(self):
        """ Set x axis as a text axis
        :return:
        """
        self.y_axis_params.update({'text_axis': True})

    def set_x_date_axis(self, major_unit, minor_unit, major_unit_type, minor_unit_type):
        """ Set x axis as a date axis.

        :param major_unit: int
        :param minor_unit: int
        :param major_unit_type: string
        :param minor_unit_type: string
        :return:
        """
        self.x_axis_params.update({
            'date_axis': True,
            'major_unit': major_unit,
            'minor_unit': minor_unit,
            'major_unit_type': major_unit_type,
            'minor_unit_type': minor_unit_type
        })

    def set_y_date_axis(self, major_unit, minor_unit, major_unit_type, minor_unit_type):
            """ Set x axis as a date axis.

            :param major_unit: int
            :param minor_unit: int
            :param major_unit_type: string
            :param minor_unit_type: string
            :return:
            """
            self.y_axis_params.update({
                'date_axis': True,
                'major_unit': major_unit,
                'minor_unit': minor_unit,
                'major_unit_type': major_unit_type,
                'minor_unit_type': minor_unit_type
            })

    def set_x_display_units(self, units, units_label=False):
        """ Set the x axis display units.

        :param units: string
        :param units_label: bool, default False
        :return:
        """
        self.chart.set_x_axis({'display_units': units, 'display_units_visible': units_label})

    def set_y_display_units(self, units, units_label=False):
        """ Set the y axis display units.

        :param units: string
        :param units_label: bool, default False
        :return:
        """
        self.chart.set_y_axis({'display_units': units, 'display_units_visible': units_label})

    def save(self, chart_sheet=None):
        """ Save the chart for chart sheet or work sheet.

        :param chart_sheet:
        :return:
        """
        self.chart.set_x_axis(self.x_axis_params)
        self.chart.set_y_axis(self.y_axis_params)
        if chart_sheet:
            chart_sheets = self.workbook.add__chartsheet(chart_sheet)
            chart_sheets.set_chart(self.chart)
        else:
            self.worksheet.insert_chart('D4', self.chart)


class ExcelChart(object):
    """ Excel Chart Class.

    """
    def __init__(self, filename):
        """

        :param filename: string
        :return:
        """
        self._filename = filename
        self._workbook = xlsxwriter.Workbook(self._filename)
        self._bold = self._workbook.add_format({'bold': 1})
        self._date_format = self._workbook.add_format({'num_format': 'yyyy-mm-dd'})
        self._charts = []

    def column(self, frame, sheet_name=None, subtype=None, data_labels=False, overlap=0, gap=150, title=None,
               font='Arial', legend=None, x_grid=False, y_grid=False, x_limit=None, y_limit=None, x_title=None,
               y_title=None, x_reverse=False, y_reverse=False, table=False, size=None, chart_sheet=None
               ):
        """ Create a Column chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param subtype: string, stacked percent_stacked, default None
        :param data_labels: bool, default False
        :param overlap: int, default 0
        :param gap: int, default 150
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param x_grid: bool, default False
        :param y_grid: bool, default False
        :param x_limit: tuple, default None
        :param y_limit: tuple, default None
        :param x_title: string, default None
        :param y_title: string, default None
        :param x_reverse: bool, default False
        :param y_reverse: bool, default False
        :param table: bool, default, False
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """

        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='column', subtype=subtype)

        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)

        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_limit(limit=x_limit)
        chart.set_y_limit(limit=y_limit)

        if x_title:
            chart.set_x_title(title=x_title, font=font)

        if y_title:
            chart.set_y_title(title=y_title, font=font)

        if x_reverse:
            chart.set_x_reverse()

        if y_reverse:
            chart.set_y_reverse()

        if table:
            chart.set_table(font=font)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def bar(self, frame, sheet_name=None, subtype=None, data_labels=False, overlap=0, gap=150, title=None,
            font='Arial', legend=None, x_grid=False, y_grid=False, x_limit=None, y_limit=None, x_title=None,
            y_title=None, x_reverse=False, y_reverse=False, table=False, size=None, chart_sheet=None
            ):
        """ Create a Bar chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param subtype: string, stacked percent_stacked, default None
        :param data_labels: bool, default False
        :param overlap: int, default 0
        :param gap: int, default 150
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param x_grid: bool, default False
        :param y_grid: bool, default False
        :param x_limit: tuple, default None
        :param y_limit: tuple, default None
        :param x_title: string, default None
        :param y_title: string, default None
        :param x_reverse: bool, default False
        :param y_reverse: bool, default False
        :param table: bool, default, False
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='bar', subtype=subtype)

        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)

        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_limit(limit=x_limit)
        chart.set_y_limit(limit=y_limit)

        if x_title:
            chart.set_x_title(title=x_title, font=font)

        if y_title:
            chart.set_y_title(title=y_title, font=font)

        if x_reverse:
            chart.set_x_reverse()

        if y_reverse:
            chart.set_y_reverse()

        if table:
            chart.set_table(font=font)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def line(self, frame, sheet_name=None, data_labels=False, overlap=0, gap=150, title=None, font='Arial',
             legend=None, x_grid=False, y_grid=False, x_limit=None, y_limit=None, x_title=None, y_title=None,
             x_reverse=False, y_reverse=False, table=False, size=None, chart_sheet=None
             ):
        """ Create a Line chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param data_labels: bool, default False
        :param title: string, default None
        :param x_title: string, default None
        :param y_title: string, default None
        :param x_grid: bool, default False
        :param y_grid: bool, default False
        :param x_reverse: bool, default False
        :param y_reverse: bool, default False
        :param x_limit: tuple, default None
        :param y_limit: tuple, default None
        :param size: tuple, default None
        :param legend: string, default None
        :param chart_sheet: string, default None
        :param font: string, default 'Arial'
        :param overlap: int, default 0
        :param gap: int, default 150
        :param table: bool, default, False
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='line')

        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)

        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_limit(limit=x_limit)
        chart.set_y_limit(limit=y_limit)

        if x_title:
            chart.set_x_title(title=x_title, font=font)

        if y_title:
            chart.set_y_title(title=y_title, font=font)

        if x_reverse:
            chart.set_x_reverse()

        if y_reverse:
            chart.set_y_reverse()

        if table:
            chart.set_table(font=font)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def pie(self, frame, sheet_name=None, data_labels=False, title=None, font='Arial', legend=None, rotation=0,
            size=None, chart_sheet=None):
        """ Create a Pie chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param data_labels: bool, default False
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param rotation: int, default 0
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='pie')
        chart.add_series(data_labels=data_labels)
        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend, font=font)
        chart.set_rotation(rotation)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def radar(self, frame, sheet_name=None, subtype=None, data_labels=False, title=None, font='Arial', legend=None,
              major_grid=True, minor_grid=False, size=None, chart_sheet=None):
        """ Create a Radar chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param subtype: string, with_markers filled, default None
        :param data_labels: bool, default False
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param major_grid: bool, default True
        :param minor_grid: bool, default False
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='radar', subtype=subtype)
        chart.add_series(data_labels=data_labels)
        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend, font=font)
        chart.set_y_grid(major=major_grid, minor=minor_grid)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def scatter(self, frame, sheet_name=None, subtype=None, data_labels=False, overlap=0, gap=150, title=None,
                font='Arial', legend=None, x_grid=False, y_grid=False, x_limit=None, y_limit=None, x_title=None,
                y_title=None, x_reverse=False, y_reverse=False, table=False, size=None, chart_sheet=None
                ):
        """ Create a Scatter chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param subtype: string, straight_with_markers straight smooth_with_markers smooth, default None
        :param data_labels: bool, default False
        :param overlap: int, default 0
        :param gap: int, default 150
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param x_grid: bool, default False
        :param y_grid: bool, default False
        :param x_limit: tuple, default None
        :param y_limit: tuple, default None
        :param x_title: string, default None
        :param y_title: string, default None
        :param x_reverse: bool, default False
        :param y_reverse: bool, default False
        :param table: bool, default, False
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='scatter', subtype=subtype)

        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)

        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_limit(limit=x_limit)
        chart.set_y_limit(limit=y_limit)

        if x_title:
            chart.set_x_title(title=x_title, font=font)

        if y_title:
            chart.set_y_title(title=y_title, font=font)

        if x_reverse:
            chart.set_x_reverse()

        if y_reverse:
            chart.set_y_reverse()

        if table:
            chart.set_table(font=font)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def area(self, frame, sheet_name=None, subtype=None, data_labels=False, overlap=0, gap=150, title=None,
             font='Arial', legend=None, x_grid=False, y_grid=False, x_limit=None, y_limit=None, x_title=None,
             y_title=None, x_reverse=False, y_reverse=False, table=False, size=None, chart_sheet=None
             ):
        """ Create a Area chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param subtype: string, stacked percent_stacked, default None
        :param data_labels: bool, default False
        :param overlap: int, default 0
        :param gap: int, default 150
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param x_grid: bool, default False
        :param y_grid: bool, default False
        :param x_limit: tuple, default None
        :param y_limit: tuple, default None
        :param x_title: string, default None
        :param y_title: string, default None
        :param x_reverse: bool, default False
        :param y_reverse: bool, default False
        :param table: bool, default, False
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='area', subtype=subtype)

        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_limit(limit=x_limit)
        chart.set_y_limit(limit=y_limit)

        if x_title:
            chart.set_x_title(title=x_title, font=font)

        if y_title:
            chart.set_y_title(title=y_title, font=font)

        if x_reverse:
            chart.set_x_reverse()

        if y_reverse:
            chart.set_y_reverse()

        if table:
            chart.set_table(font=font)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def doughnut(self, frame, sheet_name=None, data_labels=False, title=None, font='Arial', legend=None, rotation=0,
                 size=None, chart_sheet=None):
        """ Create a Doughnut chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param data_labels: bool, default False
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param rotation: int, default 0
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='doughnut')
        chart.add_series(data_labels=data_labels)
        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend, font=font)
        chart.set_rotation(rotation)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def stock(self, frame, sheet_name=None, data_labels=False, overlap=0, gap=150, title=None,
              font='Arial', legend=None, x_grid=False, y_grid=False, x_limit=None, y_limit=None, x_title=None,
              y_title=None, x_reverse=False, y_reverse=False, table=False, size=None, chart_sheet=None
              ):
        """ Create a Area chart.

        :param frame: DataFrame
        :param sheet_name: string, default None
        :param data_labels: bool, default False
        :param overlap: int, default 0
        :param gap: int, default 150
        :param title: string, default None
        :param font: string, default 'Arial'
        :param legend: string, default None
        :param x_grid: bool, default False
        :param y_grid: bool, default False
        :param x_limit: tuple, default None
        :param y_limit: tuple, default None
        :param x_title: string, default None
        :param y_title: string, default None
        :param x_reverse: bool, default False
        :param y_reverse: bool, default False
        :param table: bool, default, False
        :param size: tuple, default None
        :param chart_sheet: string, default None
        :return:
        """
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='stock', date_parser=True)

        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font=font)
        chart.set_legend(legend=legend)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_limit(limit=x_limit)
        chart.set_y_limit(limit=y_limit)

        if x_title:
            chart.set_x_title(title=x_title, font=font)

        if y_title:
            chart.set_y_title(title=y_title, font=font)

        if x_reverse:
            chart.set_x_reverse()

        if y_reverse:
            chart.set_y_reverse()

        if table:
            chart.set_table(font=font)

        if size:
            chart.set_size(width=size[0], height=size[1])

        self._charts.append((chart, chart_sheet))

        return chart

    def save(self):
        """ Save chart for Excel.

        :return:
        """
        for chart, chart_sheet in self._charts:
            chart.save(chart_sheet)
        self._workbook.close()


if __name__ == '__main__':
    bar = pd.read_excel('data/bar2.xlsx')
    # pie = pd.read_excel('data/pie.xlsx')
    lines = pd.read_excel('data/line.xlsx')
    scatter = pd.read_excel('data/scatter.xlsx')
    radar = pd.read_excel('data/radar.xlsx')
    # stock = pd.read_excel('data/stock.xlsx')

    ec = ExcelChart('chart.xlsx')

    ax = ec.bar(bar, sheet_name='bar', legend='top', y_limit=(20, 70))
    # ax.set_size()
    # ax.set_title('标题ABC123')
    # ax.set_legend('top')
    # ax.set_x_axis('标题ABC123')
    # ax.set_x_title('标题ABC123')
    # ax.set_x_label(interval_unit=2)
    # ax.set_x_tick(interval_unit=2, major_type='inside', minor_type='outside')
    # ax.set_x_interval(label_interval=2, tick_interval=2)
    # ax.set_x_limit(limit=(0.5, 5.5))
    # ax.set_x_tick_mark(major_type='inside')
    # ax.set_y_tick_mark(major_type='outside')
    # ax.set_x_title('标题ABC123')
    # ax.set_x_label(font='华文彩云')
    # ax.set_y_title('标题ABC123', font_name='华文彩云')

    # ax2 = ec.barh(bar, sheet_name='barh')
    # ax3 = ec.line(line, sheet_name='line')
    # ax3.set_x_limit(limit=(2, 7))
    # ax4 = ec.area(bar, sheet_name='area')
    # ax5 = ec.pie(pie, sheet_name='pie')
    # ax6 = ec.scatter(scatter, sheet_name='scatter')
    # ax6.set_y_grid()
    # ax7 = ec.doughnut(pie, sheet_name='doughnut')
    ax8 = ec.radar(radar, sheet_name='radar')
    # ax9 = ec.stock(stock)

    ec.save()
