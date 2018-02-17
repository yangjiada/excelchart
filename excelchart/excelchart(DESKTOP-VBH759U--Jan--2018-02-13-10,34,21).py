#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: Jan Yang
@software: PyCharm Community Edition
@time: 2017/12/26 13:14
"""


import pandas as pd
import string
import xlsxwriter


class Chart(object):
    def __init__(self, workbook, frame, sheet_name=None, chart_type=None, subtype=None, date_parser=False):
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

        if sheet_name is None:
            self.sheet_name = self.worksheet.name

        # Write data to worksheet.
        self.worksheet.write_row('A1', self.frame.columns, self.bold)

        for col in range(self.shape[1]):
            """
            self.worksheet.write_column('{}2'.format(self.uppercase[col]), self.data[:, col],
                                        self.date_format if date_parser and col == 0 else None)
            """
            self.worksheet.write_column(1, col, self.data[:, col],
                                        self.date_format if date_parser and col == 0 else None)

        # Create a Chart object.
        self.chart = self.workbook.add_chart({'type': chart_type, 'subtype': subtype})

    def add_series(self, data_labels, overlap=0, gap=150):
        """ Add one or more data series.

        :param data_labels:
        :param overlap:
        :param gap:
        :return:
        """
        """
        for num in range(1, self.shape[1]):
            self.chart.add_series({
                'name': '={}!${}$1'.format(self.sheet_name, self.uppercase[num]),
                'categories': '={}!${col}$2:${col}${row}'.format(self.sheet_name, col=self.uppercase[0],
                                                                 row=self.shape[0] + 1),
                'values': '={}!${col}$2:${col}${row}'.format(self.sheet_name, col=self.uppercase[num],
                                                             row=self.shape[0] + 1),
                'data_labels': {'value': data_labels},
                'overlap': overlap,
                'gap': gap
            })
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

    def set_size(self, width=480, height=350, x_scale=0, y_scale=0, x_offset=0, y_offset=0):
        """ Set size of the chart.

        :param width: int, default 480
        :param height: int, default 288
        :param x_scale:
        :param y_scale:
        :param x_offset:
        :param y_offset:
        :return:
        """
        if False:
            self.chart.set_size({
                'width': width,
                'height': height,
                'x_scale': x_scale,
                'y_scale': y_scale,
                'x_offset': x_offset,
                'y_offset': y_offset
            })

        self.chart.set_size({'width': width, 'height': height})

    def set_title(self, title, font_name='Arial', font_size=16, bold=True, italic=False, underline=False,
                  color='black', rotation=0, overlay=False, layout=None):
        """ Set the chart title.

        :param title:
        :param font_name:
        :param font_size:
        :param bold:
        :param italic:
        :param underline:
        :param color:
        :param rotation:
        :param overlay:
        :param layout:
        :return:
        """
        _ = layout  # occupied

        if title:
            self.chart.set_title({
                'name': title,
                'name_font': {
                    'name': font_name,
                    'size': font_size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                },
                'overlay': overlay
            })
        # else:
        #     self.chart.set_title({'none': not title})

    def set_legend(self, legend, font_name='Arial', font_size=10, bold=False, italic=False, underline=False,
                   rotation=0, color='black', delete_series=None, layout=None):
        """ Set the chart legend.

        :param legend:
        :param font_name:
        :param font_size:
        :param bold:
        :param italic:
        :param underline:
        :param rotation:
        :param color:
        :param delete_series:
        :param layout:
        :return:
        """

        _ = layout

        # top bottom left right overlay_left overlay_right none
        if legend:
            self.chart.set_legend({
                'position': legend,
                'font': {
                    'name': font_name,
                    'size': font_size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                },
                'delete_series': delete_series
            })

    def set_chart_area(self, border=False, border_color='black',  width=0.75, border_transparency=50, dash_type='solid',
                       fill=False, fill_color='white', fill_transparency=0, pattern=None, gradient=None):
        """ Set the chart area.

        :param border:
        :param border_color:
        :param width:
        :param border_transparency:
        :param dash_type:
        :param fill:
        :param fill_color:
        :param fill_transparency
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
                      fill=False, fill_color='white', fill_transparency=0, pattern=None, gradient=None):
        """ Set the plot area.

        :param border:
        :param border_color:
        :param width:
        :param border_transparency:
        :param dash_type:
        :param fill:
        :param fill_color:
        :param fill_transparency:
        :param pattern:
        :param gradient:
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
            } if fill else {'none': True}
        })

    def set_style(self, style_id):
        """ Set the chart style type.

        :param style_id:
        :return:
        """
        if style_id:
            self.chart.set_style(style_id)

    def set_table(self, table=False, horizontal=True, vertical=True, outline=True, show_keys=False, font_name='Arial',
                  font_size=10, bold=False, italic=False, underline=False, rotation=0, color='black'):
        """ Set data table.

        :param table:
        :param horizontal:
        :param vertical:
        :param outline:
        :param show_keys:
        :param font_name:
        :param font_size:
        :param bold:
        :param italic:
        :param underline:
        :param rotation:
        :param color:
        :return:
        """
        if table:
            self.chart.set_table({
                'horizontal': horizontal,
                'vertical': vertical,
                'outline': outline,
                'show_keys': show_keys,
                'font': {
                    'name': font_name,
                    'size': font_size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                }
            })

    def set_up_down_bars(self, up_down_bars, up_border_color='black', up_width=0.75, up_border_transparency=50,
                         up_dash_type='solid',
                         up_fill_color='green', up_fill_transparency=50, down_border_color='black', down_width=0.75,
                         down_border_transparency=50, down_dash_type='solid', down_fill_color='red',
                         down_fill_transparency=0):
        """ Set properties for the chart up-down bars.

        :param up_down_bars:
        :param up_border_color:
        :param up_width:
        :param up_border_transparency:
        :param up_dash_type:
        :param up_fill_color:
        :param up_fill_transparency:
        :param down_border_color:
        :param down_width:
        :param down_border_transparency:
        :param down_dash_type:
        :param down_fill_color:
        :param down_fill_transparency:
        :return:
        """

        if up_down_bars:
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

        :param color:
        :param width:
        :param dash_type:
        :param transparency:
        :return:
        """

        # solid round_dot square_dot dash dash_dot long_dash long_dash_dot long_dash_dot_dot
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

        :param color:
        :param width:
        :param dash_type:
        :param transparency:
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

        :param show:
        :return:
        """

        # gap zero span
        self.chart.show_blanks_as(show)

    def show_hidden_data(self):
        """ Display data on _charts from hidden rows or columns.

        :return:
        """
        self.show_hidden_data()

    def set_rotation(self, rotation=0):
        """ Set the Pie/Doughnut chart rotation.

        :param rotation: 0 <= rotation <= 360
        :return:
        """
        self.chart.set_rotation(rotation)

    def set_hole_size(self, size):
        """ Set the Doughnut chart hole size.

        :param size:
        :return:
        """
        # 10 <= size <= 90
        self.set_hole_size(size)

    def set_x_axis(self, font_name='Arial', font_size=10, bold=False, italic=False, underline=False, rotation=0,
                   font_color='black', num_format=None, line=True, line_color='black', width=0.75, dash_type='solid',
                   line_transparency=50, fill=False, fill_color='white', fill_transparency=0):
        self.chart.set_x_axis({
            'num_font': {
                'name': font_name,
                'size': font_size,
                'bold': bold,
                'italic': italic,
                'underline': underline or None,
                'rotation': rotation,
                'color': font_color,
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

        :param major_type: inside outside 
        :param minor_type:
        :return:
        """
        self.chart.set_x_axis({
            'major_tick_mark': major or 'none',
            'minor_tick_mark': minor or 'none'
        })

    def set_y_tick_mark(self, major='outside', minor=None):
        self.chart.set_y_axis({
            'major_tick_mark': major or 'none',
            'minor_tick_mark': minor or 'none'
        })

    def set_x_label(self, title, font_name='Arial', font_size=10, bold=True, italic=False, underline=False, rotation=0,
                    color='black'):
        """ Set the chart x label title.

        :param title:
        :param font_name:
        :param font_size:
        :param bold:
        :param italic:
        :param underline:
        :param rotation:
        :param color:
        :return:
        """

        if title:
            self.chart.set_x_axis({
                'name': title,
                'name_font': {
                    'name': font_name,
                    'size': font_size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                }
            })

    def set_y_label(self, title, font_name='Arial', font_size=10, bold=True, italic=False, underline=False, rotation=0,
                    color='black'):
        """ Set the chart y label title.

        :param title:
        :param font_name:
        :param font_size:
        :param bold:
        :param italic:
        :param underline:
        :param rotation:
        :param color:
        :return:
        """
        if title:
            self.chart.set_y_axis({
                'name': title,
                'name_font': {
                    'name': font_name,
                    'size': font_size,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline or None,
                    'rotation': rotation,
                    'color': color,
                }
            })

    def set_x_grid(self, major=False, minor=False, major_color='black', major_width=0.75, major_dash_type='solid',
                   major_transparency=50, minor_color='black', minor_width=0.75, minor_dash_type='solid',
                   minor_transparency=50):
        self.chart.set_x_axis({
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

        self.chart.set_y_axis({
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

    def set_x_reverse(self, reverse):
        if reverse:
            self.chart.set_x_axis({'reverse': reverse})

    def set_y_reverse(self, reverse):
        if reverse:
            self.chart.set_y_axis({{'reverse': reverse}})

    def set_x_lim(self, lim):
        if lim:
            self.chart.set_x_axis({'min': lim[0], 'max': lim[1]})

    def set_y_lim(self, lim):
        if lim:
            self.chart.set_y_axis({'min': lim[0], 'max': lim[1]})

    def set_x_unit(self, major, minor):
        self.chart.set_x_axis({'major_unit': major, 'minor_unit': minor})

    def set_y_unit(self, major, minor):
        self.chart.set_y_axis({'major_unit': major, 'minor_unit': minor})

    def set_x_crossing(self, category):
        self.chart.set_x_axis({'crossing': category})

    def set_y_crossing(self, category):
        self.chart.set_y_axis({'crossing': category})

    def set_x_position(self, category):
        """

        :param category: on_tick or between
        :return:
        """
        self.chart.set_x_axis({'position_axis': category})

    def set_y_position(self, category):
        self.chart.set_y_axis({'position_axis': category})

    def set_x_log(self, base):
        self.chart.set_x_axis({'log_base': base})

    def set_y_log(self, base):
        self.chart.set_y_axis({'log_base': base})

    def set_x_label_position(self, category):
        """

        :param category: next_to high low none, default next_to
        :return:
        """
        self.chart.set_x_axis({'label_position': category})

    def set_y_label_position(self, category):
        """

        :param category: next_to high low none, default next_to
        :return:
        """
        self.chart.set_y_axis({'label_position': category})

    def set_x_visible(self, visible=True):
        self.chart.set_x_axis({'visible': visible})

    def set_y_visible(self, visible=True):
        self.chart.set_y_axis({'visible': visible})

    def set_x_type(self, category):
        if category == 'date':
            self.chart.set_x_axis({'date_axis': True})
        elif category == 'text':
            self.chart.set_x_axis({'text_axis': True})

    def set_x_display_units(self, units, units_label=False):
        self.chart.set_x_axis({'display_units': units, 'display_units_visible': units_label})

    def set_y_display_units(self, units, units_label=False):
        self.chart.set_y_axis({'display_units': units, 'display_units_visible': units_label})

    def save(self, chart_sheet=None):
        if chart_sheet:
            chart_sheets = self.workbook.add__chartsheet(chart_sheet)
            chart_sheets.set_chart(self.chart)
        else:
            self.worksheet.insert_chart('D4', self.chart)


class ExcelChart(object):
    def __init__(self, filename):
        # self._filename = filename

        self._workbook = xlsxwriter.Workbook(filename)
        self._bold = self._workbook.add_format({'bold': 1})
        self._date_format = self._workbook.add_format({'num_format': 'yyyy-mm-dd'})
        self._charts = []

    def bar(self, frame, sheet_name=None,  subtype=None, data_labels=False, title=None, x_label=None, y_label=None,
            x_grid=False, y_grid=False, x_reverse=False, y_reverse=False, x_lim=None, y_lim=None, size=None,
            legend=None, chart_sheet=None, font_name='Arial', overlap=0, gap=150, table=False
            ):

        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='column', subtype=subtype)

        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)

        if size:
            chart.set_size(width=size[0], height=size[0])

        chart.set_title(title=title, font_name=font_name)
        chart.set_legend(legend=legend)
        chart.set_chart_area(border=True, fill=True)
        chart.set_plot_area(border=False, fill=True)
        chart.set_style(None)
        chart.set_table(table=table)

        chart.set_x_label(title=x_label, font_name=font_name)
        chart.set_y_label(title=y_label, font_name=font_name)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_reverse(reverse=x_reverse)
        chart.set_y_reverse(reverse=y_reverse)
        chart.set_x_lim(lim=x_lim)
        chart.set_y_lim(lim=y_lim)

        chart.set_x_axis()


        self._charts.append((chart, chart_sheet))

        return chart

    def barh(self, frame, sheet_name=None,  subtype=None, data_labels=False, title=None, x_label=None, y_label=None,
             x_grid=False, y_grid=False, x_reverse=False, y_reverse=False, x_lim=None, y_lim=None, size=None,
             legend=None, chart_sheet=None, font_name='微软雅黑', overlap=0, gap=150, table=False
             ):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='bar', subtype=subtype)
        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font_name=font_name)
        chart.set_x_label(title=x_label, font_name=font_name)
        chart.set_y_label(title=y_label, font_name=font_name)
        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)
        chart.set_x_reverse(reverse=x_reverse)
        chart.set_y_reverse(reverse=y_reverse)
        chart.set_x_lim(lim=x_lim)
        chart.set_y_lim(lim=y_lim)
        chart.set_table(table=table)
        chart.set_legend(legend=legend)
        if size:
            chart.set_size(width=size[0], height=size[0])
        # chart.save(chart_sheet=chart_sheet)
        self._charts.append((chart, chart_sheet))

        return chart

    def line(self, frame, sheet_name=None,  subtype=None, data_labels=False, title=None, x_label=None, y_label=None,
             x_grid=False, y_grid=False, x_reverse=False, y_reverse=False, x_lim=None, y_lim=None, size=None,
             legend=None, chart_sheet=None, font_name='微软雅黑', overlap=0, gap=150, table=False
             ):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='line', subtype=subtype)
        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font_name=font_name)

        # chart.set_x_axis(font_name='微软雅黑')

        chart.set_x_label(title=x_label, font_name=font_name)
        chart.set_y_label(title=y_label, font_name=font_name)
        chart.set_x_grid(major=x_grid)
        chart.set_y_grid(major=y_grid)
        chart.set_x_reverse(reverse=x_reverse)
        chart.set_y_reverse(reverse=y_reverse)
        chart.set_x_lim(lim=x_lim)
        chart.set_y_lim(lim=y_lim)
        chart.set_table(table=table)
        chart.set_legend(legend=legend)
        # chart.set_up_down_bars(False)
        # chart.set_drop_lines()
        # chart.set_high_low_lines()

        if size:
            chart.set_size(width=size[0], height=size[0])

        self._charts.append((chart, chart_sheet))

        return chart

    def pie(self, frame, sheet_name=None, title=None, chart_sheet=None, font_name='微软雅黑', subtype=None, legend=None,
            size=None, rotation=0):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='pie', subtype=subtype)
        chart.add_series(data_labels=None)
        chart.set_title(title=title, font_name=font_name)
        chart.set_legend(legend=legend)
        if size:
            chart.set_size(width=size[0], height=size[0])
        chart.set_rotation(rotation)
        self._charts.append((chart, chart_sheet))

        return chart

    def radar(self, frame, sheet_name=None, title=None, size=None, data_labels=False, subtype=None,
              chart_sheet=None, font_name='微软雅黑', legend=None):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='radar', subtype=subtype)
        chart.add_series(data_labels=data_labels)
        chart.set_title(title=title, font_name=font_name)
        chart.set_legend(legend=legend)
        if size:
            chart.set_size(width=size[0], height=size[0])
        self._charts.append((chart, chart_sheet))

        return chart

    def scatter(self, frame, sheet_name=None,  subtype=None, data_labels=False, title=None, x_label=None, y_label=None,
                x_grid=False, y_grid=False, x_reverse=False, y_reverse=False, x_lim=None, y_lim=None, size=None,
                legend=None, chart_sheet=None, font_name='微软雅黑', overlap=0, gap=150, table=False
                ):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='scatter', subtype=subtype)
        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font_name=font_name)

        chart.set_x_label(title=x_label, font_name=font_name)
        chart.set_y_label(title=y_label, font_name=font_name)

        chart.set_x_grid(major=x_grid, minor=False)
        chart.set_y_grid(major=y_grid, minor=False)

        chart.set_x_reverse(reverse=x_reverse)
        chart.set_y_reverse(reverse=y_reverse)

        chart.set_x_lim(lim=x_lim)
        chart.set_y_lim(lim=y_lim)

        chart.set_table(table=table)
        chart.set_legend(legend=legend)

        chart.set_x_ticks()
        # chart.set_y_ticks()

        if size:
            chart.set_size(width=size[0], height=size[0])
        self._charts.append((chart, chart_sheet))

        return chart

    def area(self, frame, sheet_name=None,  subtype=None, data_labels=False, title=None, x_label=None, y_label=None,
             x_grid=False, y_grid=False, x_reverse=False, y_reverse=False, x_lim=None, y_lim=None, size=None,
             legend=None, chart_sheet=None, font_name='微软雅黑', overlap=0, gap=150, table=False
             ):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='area', subtype=subtype,
                      )
        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font_name=font_name)
        chart.set_x_label(title=x_label, font_name=font_name)
        chart.set_y_label(title=y_label, font_name=font_name)
        chart.set_x_grid(major=x_grid)
        chart.set_y_grid(major=y_grid)
        chart.set_x_reverse(reverse=x_reverse)
        chart.set_y_reverse(reverse=y_reverse)
        chart.set_x_lim(lim=x_lim)
        chart.set_y_lim(lim=y_lim)
        chart.set_table(table=table)
        chart.set_legend(legend=legend)

        if size:
            chart.set_size(width=size[0], height=size[0])

        self._charts.append((chart, chart_sheet))

        return chart

    def doughnut(self, frame, sheet_name=None,  subtype=None, data_labels=False, title=None, x_label=None, y_label=None,
                 x_grid=False, y_grid=False, x_reverse=False, y_reverse=False, x_lim=None, y_lim=None, size=None,
                 legend=None, chart_sheet=None, font_name='微软雅黑', overlap=0, gap=150, table=False
                 ):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='doughnut', subtype=subtype,
                      )
        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font_name=font_name)
        chart.set_x_label(title=x_label, font_name=font_name)
        chart.set_y_label(title=y_label, font_name=font_name)
        chart.set_x_grid(major=x_grid)
        chart.set_y_grid(major=y_grid)
        chart.set_x_reverse(reverse=x_reverse)
        chart.set_y_reverse(reverse=y_reverse)
        chart.set_x_lim(lim=x_lim)
        chart.set_y_lim(lim=y_lim)
        chart.set_table(table=table)
        chart.set_legend(legend=legend)
        if size:
            chart.set_size(width=size[0], height=size[0])
        self._charts.append((chart, chart_sheet))

        return chart

    def stock(self, frame, sheet_name=None,  subtype=None, data_labels=False, title=None, x_label=None, y_label=None,
              x_grid=False, y_grid=False, x_reverse=False, y_reverse=False, x_lim=None, y_lim=None, size=None,
              legend=None, chart_sheet=None, font_name='微软雅黑', overlap=0, gap=150, table=False
              ):
        chart = Chart(self._workbook, frame=frame, sheet_name=sheet_name, chart_type='stock', subtype=subtype,
                      date_parser=True)
        chart.add_series(data_labels=data_labels, overlap=overlap, gap=gap)
        chart.set_title(title=title, font_name=font_name)
        chart.set_x_label(title=x_label, font_name=font_name)
        chart.set_y_label(title=y_label, font_name=font_name)
        chart.set_x_grid(major=x_grid)
        chart.set_y_grid(major=y_grid)
        chart.set_x_reverse(reverse=x_reverse)
        chart.set_y_reverse(reverse=y_reverse)
        chart.set_x_lim(lim=x_lim)
        chart.set_y_lim(lim=y_lim)
        chart.set_table(table=table)
        chart.set_legend(legend=legend)
        if size:
            chart.set_size(width=size[0], height=size[0])
        self._charts.append((chart, chart_sheet))

        return chart

    def save(self):
        for chart, chart_sheet in self._charts:
            chart.save(chart_sheet)
        self._workbook.close()


if __name__ == '__main__':
    bar = pd.read_excel('data/bar.xlsx')
    # pie = pd.read_excel('data/pie.xlsx')
    # line = pd.read_excel('data/line.xlsx')
    scatter = pd.read_excel('data/scatter.xlsx')
    # radar = pd.read_excel('data/radar.xlsx')
    # stock = pd.read_excel('data/stock.xlsx')

    ec = ExcelChart('chart.xlsx')

    ax = ec.bar(bar, sheet_name='bar', legend='top', title='bar title')
    ax.set_y_tick_mark()

    # ax2 = ec.barh(bar, sheet_name='barh')
    # ax3 = ec.line(line, sheet_name='line')
    # ax4 = ec.area(bar, sheet_name='area')
    # ax5 = ec.pie(pie, sheet_name='pie')
    # ax6 = ec.scatter(scatter, sheet_name='scatter')
    # ax6.set_y_grid()
    # ax7 = ec.doughnut(pie, sheet_name='doughnut')
    # ax8 = ec.radar(radar, sheet_name='radar')
    # ax9 = ec.stock(stock)

    ec.save()
    print()
