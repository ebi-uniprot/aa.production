from xls_util import XlsUtil
from readData import Report, Section
from analysisData import DiffReport, DiffSection

try:
    import xlsxwriter
except ImportError:
    print('\nThere was no xlswriter module installed. You can install it with:\npip install xlsxwriter')
    sys.exit(1)

# Worksheet class for writing individual reports and differetial report
class Worksheet:
    def __init__(self, workbook, name):
        self.worksheet = workbook.add_worksheet(name)
        self.row = 0
        self.format1 = workbook.add_format({'bold': True, 'underline': True, 'align': 'center', 'center_across': True,
                                            'font_name': 'Arial', 'font_size': 10})
        self.format2 = workbook.add_format({'bg_color': 'orange', 'font_name': 'Arial', 'font_size': 10})
        self.format3 = workbook.add_format({'bg_color': 'green', 'font_name': 'Arial', 'font_size': 10})
        self.format4 = workbook.add_format({'bg_color': 'blue', 'font_name': 'Arial', 'font_size': 10})
        self.format5 = workbook.add_format({'num_format': '0.00%', 'font_name': 'Arial', 'font_size': 10})
        self.format6 = workbook.add_format({'font_name': 'Arial', 'font_size': 10})

    def print_headers(self, name, headers):
        # write headers
        self.worksheet.write(self.row, 0, name, self.format1)

        # according to the length of the headers list, write the headers in the according column
        col = 1
        while True:
            col = self.write_headers(col, headers, self.format1)
            if col > 12:
                break

    def append(self, s):
        self.worksheet.write(self.row, 0, s.name, self.format1)

        # from the next row, write the data
        self.write_headers(1, s.headers, self.format1)
        self.row += 1

        for (name, numbers) in s.data:
            self.worksheet.write(self.row, 0, name, self.format6)
            self.write_numbers(1, numbers, self.format6)
            self.row += 1
        self.row += 1

    def write_headers(self, col, h, f):
        for c in range(0, len(h)):
            if len(h) == 1:
                self.worksheet.write(self.row, col + 1, h[c], f)
            else:
                self.worksheet.write(self.row, col + c, h[c], f)
                for i in (0, (3 - len(h) + 1)):
                    self.worksheet.write(self.row, col + i, None)
        return (col + 3)

    def write_numbers(self, col, n, f):
        is_num_list(n)
        for c in range(0, len(n)):
            if len(n) == 1:
                self.worksheet.write(self.row, col + 1, n[c], f)
            else:
                self.worksheet.write(self.row, col + c, n[c], f)
                for i in (0, (3 - len(n) + 1)):
                    self.worksheet.write(self.row, col + i, None)
        return (col + 3)

    def appendDiff(self, diffSec, r1, r2):
        # TODO merge the cells for main header
        self.worksheet.merge_range('B1:D1', r1.name, self.format1)
        self.worksheet.merge_range('E1:G1', r1.name, self.format1)
        self.worksheet.merge_range('H1:J1',
                                "increase {} --> {}, abs".format(r1.name, r2.name),
                                self.format1)
        self.worksheet.merge_range('K1:M1',
                                "increase {} --> {}, %".format(r1.name, r2.name),
                                self.format1)
        self.row += 1
        (name, headers, diffData) = diffSec
        self.print_headers(name, headers)
        self.row += 1

        for line in diffData.diffSec:
            col = 0
            # when there is a difference in name, only write one set of data
            if len(line) == 2:
                (lineName, nb) = line
                self.worksheet.write(self.row, col, lineName, self.format6)
                col += 1
                col = self.write_numbers(col, nb, self.format6)

            # write two sets of data with the same name
            elif len(line) == 3:
                (lineName, nb1, nb2) = line
                self.worksheet.write(self.row, col, lineName, self.format6)
                col += 1
                col = self.write_numbers(col, nb1, self.format6)
                col = self.write_numbers(col, nb2, self.format6)
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

                col = self.write_numbers(col, v, self.format6)
                col = self.write_numbers(col, p, self.format5)

            else:
                print("error")

            self.row += 1
        self.row += 1

        # conditional formatting the percentages columns
        conRange = 'K3:M105'
        self.worksheet.conditional_format(conRange, {'type':     'cell',
                                                     'criteria': '<',
                                                     'value':     0,
                                                     'format':    self.format2})

        self.worksheet.conditional_format(conRange, {'type':     'cell',
                                                     'criteria': 'between',
                                                     'minimum':   0.05,
                                                     'maximum':   0.10,
                                                     'format':    self.format3})
        self.worksheet.conditional_format(conRange, {'type':     'cell',
                                                     'criteria': '>',
                                                     'value':     0.10,
                                                     'format':    self.format4})
        # writing legend
        self.worksheet.write(3, 14, 'Legend', self.format1)
        self.worksheet.write(4, 14, 'cutoff values (change to alter colouring)', self.format1)
        self.worksheet.write(5, 14, 'decrease:  0%', self.format2)
        self.worksheet.write(6, 14, 'increase:  5%', self.format3)
        self.worksheet.write(7, 14, 'big increase:  10%', self.format4)

# Writer class to open a workbook and write in the worksheets.
class Writer:
    def __init__(self, filename):
        self.workbook = xlsxwriter.Workbook(filename)

    # take a report class and write to a worksheet
    def writeReport(self, report):
        worksheet = Worksheet(self.workbook, report.name)
        for s in report.listOfSections:
            worksheet.append(s)

    def writeDiffReport(self, r1, r2):
        xUtil = XlsUtil()
        print("r1.name: {}; r2.name: {}".format(r1.name, r2.name))
        worksheet = Worksheet(self.workbook,
                              "compare-{}".format(xUtil.generate_deviations_sheet_name(r1.name, r2.name)))
        diffR = DiffReport(r1, r2)

        for diffSec in diffR.diffSec:
            worksheet.appendDiff(diffSec, r1, r2)

    def close(self):
        self.workbook.close()

def is_num_list(l):
    if len(l) == 0:
        return
    for i in l:
        if not (type(i) == type(int()) or type(i) == type(float())):
            print("is_num_list: %", l)
            raise ValueError
