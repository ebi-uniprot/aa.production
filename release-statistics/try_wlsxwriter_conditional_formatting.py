"""
produce a sample Excel file whith conditional formatting in it.
Used to try xlswriter capabilities.
Assumes libreoffice is installed and can be found in $PATH
"""

import xlsxwriter
from subprocess import call
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol


OUT_XLSX_FILE = 'demo.xlsx'


class CellFormats:
    def __init__(self, workbook):
        self.fmt_header = workbook.add_format({'bold': True, 'align': 'center'})
        self.fmt_percent = workbook.add_format({'num_format': '0.00%'})
        self.fmt_decrease = workbook.add_format(
            {'bg_color': '#ff6633', 'font_name': 'Arial', 'font_size': 10})
        self.fmt_increase_small = workbook.add_format(
            {'bg_color': '#00ff00', 'font_name': 'Arial', 'font_size': 10})
        self.fmt_increase_big = workbook.add_format(
            {'bg_color': '#008080', 'font_name': 'Arial', 'font_size': 10})


def write_data_single_col(worksheet):
    worksheet.set_column('A:A', 15)
    worksheet.write_string('A1', 'abs values', cell_formats.fmt_header)
    worksheet.write_number('A2', -5.0)
    worksheet.write_number('A3', 0.0)
    worksheet.write_number('A4', 1.0)
    worksheet.write_number('A5', 15.0)
    worksheet.write_string('B1', '%', cell_formats.fmt_header)
    worksheet.write_number('B2', -.01, cell_formats.fmt_percent)
    worksheet.write_number('B3', 0.0, cell_formats.fmt_percent)
    worksheet.write_number('B4', .06, cell_formats.fmt_percent)
    worksheet.write_number('B5', .45, cell_formats.fmt_percent)

    (row, last_col) = xl_cell_to_rowcol('B5')
    return 'A2:A5', 'B2', last_col


def write_data_multi_col(worksheet):
    worksheet.merge_range(0, 0, 0, 1, 'abs values', cell_formats.fmt_header)
    worksheet.write_number('A2', -5.0)
    worksheet.write_number('A3', 0.0)
    worksheet.write_number('A4', 2.0)
    worksheet.write_number('A5', 15.0)
    worksheet.write_number('B2', -2.0)
    worksheet.write_number('B3', 0.01)
    worksheet.write_number('B4', 3.0)
    worksheet.write_number('B5', 3.3)

    worksheet.merge_range(0, 2, 0, 3, 'rel increase', cell_formats.fmt_header)
    data_cell = 'A2'
    (data_cell_row, data_cell_col) = xl_cell_to_rowcol(data_cell)
    targ_cell = 'C2'
    (targ_cell_row, targ_cell_col) = xl_cell_to_rowcol(targ_cell)

    for rows in range(0, 4):
        for cols in range(0, 2):
            src_data_cell = xl_rowcol_to_cell(data_cell_row + rows, data_cell_col + cols)
            worksheet.write_formula(targ_cell_row + rows, targ_cell_col + cols,
                                    '={}*2/100'.format(src_data_cell), cell_formats.fmt_percent)

    (row, last_col) = xl_cell_to_rowcol('D5')
    return 'A2:B5', 'C2', last_col


def write_thresholds_table(worksheet, col):
    worksheet.set_column(col, col, 15)
    worksheet.write_string(0, col, 'thresholds', cell_formats.fmt_header)
    worksheet.write_string(1, col, 'decrease')
    worksheet.write_string(2, col, 'increase small')
    worksheet.write_string(3, col, 'increase big')

    worksheet.write_number(1, col+1, .0, cell_formats.fmt_percent)
    worksheet.write_number(2, col+1, .05, cell_formats.fmt_percent)
    worksheet.write_number(3, col+1, .1, cell_formats.fmt_percent)

    cell_decrease = xl_rowcol_to_cell(1, col+1, row_abs=True, col_abs=True)
    cell_increase_small = xl_rowcol_to_cell(2, col+1, row_abs=True, col_abs=True)
    cell_increase_big = xl_rowcol_to_cell(3, col+1, row_abs=True, col_abs=True)

    return cell_decrease, cell_increase_small, cell_increase_big


workbook = None


def add_conditional_formatting(worksheet, cond_fmt_range, cell_decrease, cell_increase_small, cell_increase_big):
    worksheet.conditional_format(cond_fmt_range, {'type': 'cell',
                                                  'criteria': '<',
                                                  'value' : cell_decrease,
                                                  'format': cell_formats.fmt_decrease})
    worksheet.conditional_format(cond_fmt_range, {'type': 'formula',
                                                  'criteria': '={}>={}'.format(rel_first_cell, cell_increase_big),
                                                  'format': cell_formats.fmt_increase_big})
    worksheet.conditional_format(cond_fmt_range, {'type': 'formula',
                                                  'criteria': '={}>={}'.format(rel_first_cell, cell_increase_small),
                                                  'format': cell_formats.fmt_increase_small})


try:
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(OUT_XLSX_FILE)

    cell_formats = CellFormats(workbook)

    # write SINGLE-columned data sheet:
    worksheet_single_col = workbook.add_worksheet(name='single-col-conditionals')
    (range_abs, rel_first_cell, col) = write_data_single_col(worksheet_single_col)

    (cell_decrease, cell_increase_small, cell_increase_big) = \
        write_thresholds_table(worksheet_single_col, col + 2)

    add_conditional_formatting(worksheet_single_col, range_abs, cell_decrease, cell_increase_small, cell_increase_big)

    # write MULTI-columned data sheet:
    worksheet_multi_col = workbook.add_worksheet(name='conditionals-multi-col')
    (range_abs_multicol, rel_first_cell, col_multicol) = write_data_multi_col(worksheet_multi_col)

    (cell_decrease, cell_increase_small, cell_increase_big) = \
        write_thresholds_table(worksheet_multi_col, col_multicol + 2)

    add_conditional_formatting(worksheet_multi_col, range_abs_multicol,
                               cell_decrease, cell_increase_small, cell_increase_big)

finally:
    if workbook is not None:
        workbook.close()

call(['libreoffice', '{}'.format(OUT_XLSX_FILE)])
