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

    def print_headers_in_deviation_report(self, name, headers):
        # write headers
        self.worksheet.write(self.row, 0, name, self.format['Header'])
        col = 1
        while True:
            col = self.write_headers(col, headers, self.format['Header'])
            if col > 12:
                break

    def write_global_formula(self, s, r, f):
        pass

    def append(self, s):
        self.worksheet.write(self.row, 0, s.name, self.format['Header'])
        if s.is_footer == False:
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
                numberCell = self.write_1num_cn(self.row, 2, number[0], self.format['Num'])
                numberCells.append(numberCell)
                self.row += 1

            tremblCell = numberCells[-1]
            for c in numberCells:
                (row, col) = xl_cell_to_rowcol(c)
                writingCell = xl_rowcol_to_cell(row, col + 1)
                formulaGlobal = '={}/{}'.format(c, tremblCell)
                self.worksheet.write_formula(writingCell, formulaGlobal, self.format['Percent'])

        self.row += 1

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
                    col += 1
                    col = self.write_numbers(col, nb1, self.format['Num'])
                    col = self.write_numbers(col, nb2, self.format['Num'])
                    v = []
                    p = []
                    for i in range(0, len(nb1)):
                        diffVal = int(nb1[i]) - int(nb2[i])
                        v.append(diffVal)
                        if int(nb2[i]) == 0:
                            p.append(0.0)
                        else:
                            #diffPer = "{:.2%}".format(diffVal / int(nb2[i]))
                            diffPer = diffVal / int(nb2[i])
                            p.append(diffPer)

                    col = self.write_numbers(col, v, self.format['Num'])
                    col = self.write_numbers(col, p, self.format['Percent'])

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
            for line in diffData.diffSec:
                (lineName, nb1, nb2) = line
                col = 0
                self.worksheet.write(self.row, col, lineName, self.format['Num'])
                col += 1
                col = self.write_numbers(col, nb1, self.format['Num'])
                col = self.write_numbers(col, nb2, self.format['Num'])
                v = []
                p = []
                for i in range(0, len(nb1)):
                    diffVal = int(nb1[i]) - int(nb2[i])
                    v.append(diffVal)
                    if int(nb2[i]) == 0:
                        p.append(0.0)
                    else:
                        diffPer = diffVal / int(nb2[i])
                        p.append(diffPer)

                col = self.write_numbers(col, v, self.format['Num'])
                col = self.write_numbers(col, p, self.format['Percent'])

                self.row += 1


        # conditional formatting the percentages columns
        conRange = 'K3:M119'
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
        # writing legend
        legend = 'cutoff values (change to alter colouring)'
        self.worksheet.set_column(15, 16, len(legend))
        self.worksheet.merge_range('O4:P4', 'Legend', self.format['Header'])
        self.worksheet.merge_range('O5:P5', legend, self.format['Header'])
        self.worksheet.merge_range('O6:P6', 'decrease:  0%', self.format['Diff_decrease'])
        self.worksheet.merge_range('O7:P7', 'increase:  5%', self.format['Diff_increase_small'])
        self.worksheet.merge_range('O8:P8', 'big increase:  10%', self.format['Diff_increase_big'])

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
        self.format['Diff_decrease'] = self.workbook.add_format({'bg_color': 'orange', 'font_name': 'Arial', 'font_size': 10})
        self.format['Diff_increase_small'] = self.workbook.add_format({'bg_color': 'green', 'font_name': 'Arial', 'font_size': 10})
        self.format['Diff_increase_big'] = self.workbook.add_format({'bg_color': 'blue', 'font_name': 'Arial', 'font_size': 10})
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
