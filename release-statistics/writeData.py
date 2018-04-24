from xls_util import XlsUtil
from readData import Report, Section
from analysisData import DiffReport, DiffSection

try:
    import xlsxwriter
except ImportError:
    print('\nThere was no xlswriter module installed. You can install it with:\npip install xlsxwriter')
    sys.exit(1)

from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol


# Worksheet class for writing individual reports and deviation report
class Worksheet:
    def __init__(self, workbook, formatting, name):
        self.worksheet = workbook.add_worksheet(name)
        self.format = formatting
        self.row = 0

    def freeze_panes(self, r, c):
        self.worksheet.freeze_panes(r, c)

    def write_headers(self, col, h, f):
        # according to the length of the headers list, write the headers in the according column
        # the maximum length of the header list is 3 as in (predictions, entries, rules)
        if len(h) == 1:
            self.write_headers_(col + 1, h, f)
        else:
            self.write_data_padding(3, col, h, f)
        return (col + 3)

    def write_headers_(self, col, h, f):
        for c in range(0, len(h)):
            self.worksheet.write(self.row, col + c, h[c], f)
        return (col + len(h))

    def write_data_padding(self, padding, col, d, f):
        after_writing = self.write_headers_(col, d, f)
        if len(d) < padding:
            for i in range(0, padding - len(d)):
                self.worksheet.write(self.row, after_writing + i, None)
            return col + padding
        else:
            return after_writing

    def write_numbers(self, col, n, f):
        if len(n) == 1:
            self.write_numbers_(col + 1, n, f)
        else:
            self.write_data_padding(3, col, n, f)
        return (col + 3)

    def write_numbers_(self, col, n, f):
        for c in range(0, len(n)):
            self.worksheet.write_number(self.row, col + c, n[c], f)
        return (col + len(n))

    def write_footer_headers(self, col, fh, f):
        while True:
            for h in range(0, len(fh)):
                self.worksheet.write(self.row, col + 1, fh[h], f)
                col += 1
            col += 1
            if col > 12:
                break

    def write_1num_cn(self, row, col, n, f):
        self.worksheet.write_number(row, col, n, f)
        return xl_rowcol_to_cell(row, col)

    def write_numList_cn(self, row, col, nl, f):
        numberCells = []
        if len(nl) == 1:
            oneCell = self.write_1num_cn(row, col + 1, nl[0], f)
            numberCells.append(oneCell)
        else:
            for n in nl:
                oneCell = self.write_1num_cn(row, col, n, f)
                col += 1
                numberCells.append(oneCell)
            if len(numberCells) < 3:
                self.worksheet.write(row, col + 1, None)
        return numberCells

    def write_num_padding(self, padding, col, d):
        if len(d) < padding:
            for i in range(0, padding - len(d)):
                self.worksheet.write(self.row, col + len(d) + i, None)

    def print_headers_in_deviation_report(self, name, headers):
        # write headers
        self.worksheet.write(self.row, 0, name, self.format['Header'])
        col = 1
        while True:
            col = self.write_headers(col, headers, self.format['Header'])
            if col > 12:
                break

    def write_global_formula(self, c, tremblCell):
        (row, col) = xl_cell_to_rowcol(c)
        writingCell = xl_rowcol_to_cell(row, col + 1)
        formulaGlobal = '={}/{}'.format(c, tremblCell)
        self.worksheet.write_formula(writingCell, formulaGlobal, self.format['Percent'])
        return writingCell

    def write_deviation_formula_abs(self, col, c1, c2, f):
        (row, col) = xl_cell_to_rowcol(c2)
        writingCell = xl_rowcol_to_cell(row, col + 3)
        formulaValDiff = '=IF(AND({}=0, {}=0), 0, {}-{})'.format(c1, c2, c1, c2)
        self.worksheet.write_formula(writingCell, formulaValDiff, f)

    def write_deviation_formula_per(self, col, c1, c2):
        (row, col) = xl_cell_to_rowcol(c2)
        writingCell = xl_rowcol_to_cell(row, col + 6)
        formulaValDiff = '=IF({}=0, 0, ({}-{})/{})'.format(c2, c1, c2, c2)
        self.worksheet.write_formula(writingCell, formulaValDiff, self.format['Percent'])

    def append(self, s):
        self.worksheet.write(self.row, 0, s.name, self.format['Header'])
        if not s.is_footer:
            # from the next row, write the data
            self.write_headers(1, s.headers, self.format['Header'])
            self.row += 1

            for (name, numbers) in s.data:
                self.worksheet.write(self.row, 0, name, self.format['Num'])
                self.write_numbers(1, numbers, self.format['Num'])
                self.row += 1
        else:
            self.worksheet.merge_range(self.row, 0, self.row, 6, s.name, self.format['Header'])
            self.row += 1
            self.worksheet.merge_range(self.row, 0, self.row, 2, s.longHeader, self.format['Header'])
            self.row += 1

            numberCells = []
            for (name, number) in s.data:
                self.worksheet.write(self.row, 0, name, self.format['Num'])
                oneCell = self.write_1num_cn(self.row, 2, number[0], self.format['Num'])
                numberCells.append(oneCell)
                self.row += 1

            trembl_entries_cell = self.fix_row(numberCells[-1])
            for c in numberCells:
                self.write_global_formula(c, trembl_entries_cell)

        self.row += 1

    def fix_row(self, cell_name):
        # "fixes" row number with a dollar sign, for formula references
        (row, col) = xl_cell_to_rowcol(cell_name)
        return xl_rowcol_to_cell(row, col, row_abs=True)

    def appendDiff(self, diffSec, r1, r2):
        # merge the cells for main header
        self.worksheet.merge_range('B1:D1', r1.name, self.format['Header'])
        self.worksheet.merge_range('E1:G1', r2.name, self.format['Header'])
        self.worksheet.merge_range('H1:J1',
                                "increase {} --> {}, abs".format(r1.name, r2.name),
                                   self.format['Header'])
        self.worksheet.merge_range('K1:M1',
                                "increase {} --> {}, %".format(r1.name, r2.name),
                                   self.format['Header'])
        self.row += 1
        if len(diffSec) != 4:
            (name, headers, diffData) = diffSec
            self.print_headers_in_deviation_report(name, headers)
            self.row += 1

            for line in diffData.diffSec:
                col = 0
                # when there is a difference in name, only write one set of data
                if len(line) == 2:
                    (lineName, nb) = line
                    self.worksheet.write(self.row, col, lineName, self.format['Header'])
                    col += 1
                    col = self.write_numbers(col, nb, self.format['Num'])

                # write two sets of data with the same name
                elif len(line) == 3:
                    (lineName, nb1, nb2) = line
                    self.worksheet.write(self.row, col, lineName, self.format['Num'])
                    numberCells1 = []
                    numberCells2 = []
                    col += 1
                    numberCells1 = self.write_numList_cn(self.row, col, nb1, self.format['Num'])
                    col += 3
                    numberCells2 = self.write_numList_cn(self.row, col, nb2, self.format['Num'])
                    col += 3
                    for i in range(0, len(numberCells1)):
                        self.write_deviation_formula_abs(col, numberCells1[i], numberCells2[i], self.format['Num'])
                        if numberCells2[i] != 0:
                            self.write_deviation_formula_per(col, numberCells1[i], numberCells2[i])
                        else:
                            self.worksheet.write(self.row, col, 0, self.format['Num'])

                else:
                    print("error")

                self.row += 1

        else:
            (name, longHeader, headers, diffData) = diffSec
            self.worksheet.merge_range(self.row, 0, self.row, 12, name, self.format['Header'])
            self.row += 1
            self.worksheet.write(self.row, 0, longHeader, self.format['Header'])
            self.write_footer_headers(1, headers, self.format['Header'])
            self.row += 1
            numberCells1 = []
            numberCells2 = []
            for line in diffData.diffSec:
                (lineName, nb1, nb2) = line
                col = 0
                self.worksheet.write(self.row, col, lineName, self.format['Num'])
                # start to write the entries at column 2, to line up with 'entries' column
                col += 2
                oneCell1 = self.write_1num_cn(self.row, col, nb1[0], self.format['Num'])
                numberCells1.append(oneCell1)
                col += 3
                oneCell2 = self.write_1num_cn(self.row, col, nb2[0], self.format['Num'])
                numberCells2.append(oneCell2)
                col += 3
                self.row += 1

            trembl_entries_cell_1 = self.fix_row(numberCells1[-1])
            formula_cell_1 = []
            formula_cell_2 = []
            for c1 in numberCells1:
                cell1 = self.write_global_formula(c1, trembl_entries_cell_1)
                formula_cell_1.append(cell1)
            trembl_entries_cell_2 = self.fix_row(numberCells2[-1])
            for c2 in numberCells2:
                cell2 = self.write_global_formula(c2, trembl_entries_cell_2)
                formula_cell_2.append(cell2)

            for i in range(0, len(numberCells1)):
                self.write_deviation_formula_abs(col, numberCells1[i], numberCells2[i], self.format['Num'])
                self.write_deviation_formula_abs(col + 1, formula_cell_1[i], formula_cell_2[i], self.format['Percent'])
                if numberCells2[i] != 0:
                    self.write_deviation_formula_per(col, numberCells1[i], numberCells2[i])
                    self.write_deviation_formula_per(col + 1, formula_cell_1[i], formula_cell_2[i])
                else:
                    self.worksheet.write(self.row, col, 0, self.format['Num'])

        # writing legend
        legend = 'Cutoff values (change to alter colouring)'
        self.worksheet.set_column(15, 16, len(legend))
        self.worksheet.merge_range('O4:P4', 'Legend', self.format['Header'])
        self.worksheet.merge_range('O5:P5', legend, self.format['Header'])
        self.worksheet.write('O6', 'decrease: ', self.format['Diff_decrease'])
        zero_value = self.worksheet.write('P6', 0, self.format['Diff_decrease'])
        self.worksheet.merge_range('O7:P7', 'increase:  5%', self.format['Diff_increase_small'])
        self.worksheet.merge_range('O8:P8', 'big increase:  10%', self.format['Diff_increase_big'])

        # conditional formatting the percentages columns
        conRange = 'K3:M113'
        self.worksheet.conditional_format(conRange, {'type':     'cell',
                                                     'criteria': '<',
                                                     'value':     0,
                                                     'format':    self.format['Diff_decrease']})

        self.worksheet.conditional_format(conRange, {'type':     'cell',
                                                     'criteria': 'between',
                                                     'minimum':   0.05,
                                                     'maximum':   0.10,
                                                     'format':    self.format['Diff_increase_small']})
        self.worksheet.conditional_format(conRange, {'type':     'cell',
                                                     'criteria': '>',
                                                     'value':     0.10,
                                                     'format':    self.format['Diff_increase_big']})

    def set_column_width(self, r):
        listOfMaxNamesLength = []

        for s in r.sections:
            maxLength = 0
            if s.is_footer == False:
                for (name, numbers) in s.data:
                    # find the max length of names in a section
                    if len(name) > maxLength:
                        maxLength = len(name)
                listOfMaxNamesLength.append(maxLength)
        # add another 2 so that the length of column cover the whole text length
        self.worksheet.set_column(0, 0, max(listOfMaxNamesLength) + 2)
        # set the following columns to a width of 10 (set to column 12 for the deviation report)
        self.worksheet.set_column(1, 12, 10)

# Writer class to open a workbook and write in the worksheets.
class Writer:
    def __init__(self, filename):
        self.workbook = xlsxwriter.Workbook(filename)
        self.format = {}
        self.format['Header'] = self.workbook.add_format({'bold': True, 'underline': True, 'align': 'center', 'center_across': True,
                                                          'font_name': 'Arial', 'font_size': 10})
        self.format['Diff_decrease'] = self.workbook.add_format({'bg_color': '#ff6633', 'font_name': 'Arial', 'font_size': 10, 'num_format': '0.00%'})
        self.format['Diff_increase_small'] = self.workbook.add_format({'bg_color': '#00ff00', 'font_name': 'Arial', 'font_size': 10})
        self.format['Diff_increase_big'] = self.workbook.add_format({'bg_color': '#008080', 'font_name': 'Arial', 'font_size': 10})
        self.format['Percent'] = self.workbook.add_format({'num_format': '0.00%', 'font_name': 'Arial', 'font_size': 10})
        self.format['Num'] = self.workbook.add_format({'num_format': '#,###', 'font_name': 'Arial', 'font_size': 10})

    # take a report class and write to a worksheet
    def writeReport(self, report):
        worksheet = Worksheet(self.workbook, self.format, report.name)
        for s in report.sections:
            worksheet.append(s)
        worksheet.set_column_width(report)
        # freeze pane on the top row and left column
        worksheet.freeze_panes(1, 1)

    def writeDiffReport(self, r1, r2):
        xUtil = XlsUtil()
        print("r1.name: {}; r2.name: {}".format(r1.name, r2.name))
        worksheet = Worksheet(self.workbook, self.format,
                              "compare-{}".format(xUtil.generate_deviations_sheet_name(r1.name, r2.name)))
        worksheet.freeze_panes(1, 1)
        # set column width with r1 as only one set of lineNames to be compared
        worksheet.set_column_width(r1)
        diffR = DiffReport(r1, r2)

        for diffSec in diffR.diffSec:
            worksheet.appendDiff(diffSec, r1, r2)


    def close(self):
        self.workbook.close()
