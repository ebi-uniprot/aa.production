#!/usr/bin/env python

"""generateReleaseStatisticsDeviationReport.py: Generates 'deviation' report showing how much each AA Production Statistics number deviated from previous release."""
__author__ = "Cherry Hanquez"
__copyright__ = "Copyright 2018, The European Bioinformatics Institute (EMBL-EBI)"
__credits__ = ["Cherry Hanquez"]
__status__ = "Development"

import sys
import argparse
import os.path
import re

try:
    import xlsxwriter
except ImportError:
    print('\nThere was no xlswriter module installed. You can install it with:\npip install xlsxwriter')
    sys.exit(1)

def is_num_list(l):
    if len(l) == 0:
        return
    for i in l:
        if not (type(i) == type(int()) or type(i) == type(float())):
            print("is_num_list: %", l)
            raise ValueError

# Report class for opening a file and return a list of sections
class Report:
    def __init__(self, filename):
        self.name = pathNameCleanUp(filename)
        self.listOfSections = []
        with open(filename, 'rt') as in_file:
            # skip first 6 lines (if an extra empty line is added when generating the report)
            for i in range(1,6):
                in_file.readline()

            # read data
            while True:
                s = parseSection(in_file)
                if s is None:
                    break
                #print("writing section " + s.name)
                self.listOfSections.append(s)

            in_file.close()

# create a class for a typical section which contains header and data parts
class Section:
    def __init__(self, name):
        self.name = name
        self.headers = []
        self.data = []
    def append(self, line):
        x = line.split()
        data_start = 0
        # treat the special case (those ends with "entries" in the file), mark
        # it in such a way that data (aka number) starts after the name of system
        while not x[data_start].endswith(":") and not x[data_start].isdecimal():
            data_start += 1

        # read the header once
        if len(self.headers) == 0:
            for i in x[data_start:]:
                if not i.isdecimal():
                    self.headers.append(i)

        # data is in a format of list of tuples, each of which contains two lists of strings
        numbers = []
        for i in x[data_start:]:
            if i.isdecimal():
                numbers.append(float(i))
        self.data.append((" ".join(x[:data_start]), numbers))

class DiffReport:
    def __init__(self, report1, report2):
        self.diffSec = []

        for sec1 in report1.listOfSections:
            for sec2 in report2.listOfSections:
                if sec1.name == sec2.name:
                    self.diffSec.append((sec1.name, sec1.headers, DiffSection(sec1.data, sec2.data)))

class DiffSection:
    def __init__(self, d1, d2):
        self.diffSec = []

        i = 0
        while i < len(d1):
            (lineName1, nb1) = d1[i]
            (lineName2, nb2) = d2[i]
            is_num_list(nb1)
            is_num_list(nb2)
            if lineName1 == lineName2:
                self.diffSec.append((lineName1, nb1, nb2))
            else:
                self.diffSec.append(d1[i])
                self.diffSec.append(d2[i])
            i += 1

# Worksheet class for writing individual reports and differetial report
class Worksheet:
    def __init__(self, workbook, name):
        self.worksheet = workbook.add_worksheet(name)
        self.row = 0
        self.format1 = workbook.add_format({'bold': True, 'underline': True, 'align': 'center'})
        self.format2 = workbook.add_format({'bg_color': 'orange'})
        self.format3 = workbook.add_format({'bg_color': 'green'})
        self.format4 = workbook.add_format({'bg_color': 'blue'})
        self.format5 = workbook.add_format({'num_format': '0.00%'})

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
        #self.print_headers(s.name, s.headers)
        self.worksheet.write(self.row, 0, s.name, self.format1)
        # for col in range(0, len(s.headers)):
        #     self.worksheet.write(self.row, col + 1, s.headers[col], self.format1)
        # from the next row, write the data
        self.write_headers(1, s.headers, self.format1)
        self.row += 1
        # for n in s.data:
        #     for col in range(0, len(n)):
        #         self.worksheet.write(self.row, col, int(n[col]))
        #     self.row += 1
        # self.row += 1

        for (name, numbers) in s.data:
            self.worksheet.write(self.row, 0, name)
            self.write_numbers(1, numbers, None)
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

    def appendDiff(self, diffSec):
        self.worksheet.write(0, 1, "Current Stat", self.format1)
        self.worksheet.write(0, 4, "Previous Stat", self.format1)
        self.worksheet.write(0, 7, "increase abs", self.format1)
        self.worksheet.write(0, 10, "increase %", self.format1)
        self.row += 1
        (name, headers, diffData) = diffSec
        self.print_headers(name, headers)
        self.row += 1
        #self.diffV = []

        for line in diffData.diffSec:
            col = 0
            # when there is a difference in name, only write one set of data
            if len(line) == 2:
                (lineName, nb) = line
                self.worksheet.write(self.row, col, lineName)
                col += 1
                col = self.write_numbers(col, nb, None)

            # write two sets of data with the same name
            elif len(line) == 3:
                (lineName, nb1, nb2) = line
                self.worksheet.write(self.row, col, lineName)
                col += 1
                col = self.write_numbers(col, nb1, None)
                col = self.write_numbers(col, nb2, None)
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

                col = self.write_numbers(col, v, None)
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
        xUtil = XlsUtil
        print("r1.name: {}; r2.name: {}".format(r1.name, r2.name))
        worksheet = Worksheet(self.workbook,
                              "compare-{}".format(xUtil.generateDeviationsSheetName(xUtil, r1.name, r2.name)))
        diffR = DiffReport(r1, r2)

        for diffSec in diffR.diffSec:
            worksheet.appendDiff(diffSec)

    def close(self):
        self.workbook.close()

# separate the file into sections whereas an empty line
def parseSection(in_file):
    dataLines = list()
    while 1:
        line = in_file.readline()
        if not line or line == "\n":
            break
        dataLines.append(line)

    if len(dataLines) == 0:
        return None

    s = Section(dataLines[0])
    for l in dataLines[1:]:
        s.append(l)
    return s

# convert data from [ ( [String], [String] ) ] -> [ [ String ] ]
# def convertList(listOfData):
#     l = []
#     for (name, numbers) in listOfData:
#         #numbers.insert(0, name)
#         l.append([name] + numbers)
#     return l

# Clean-up of path name so it doesn't contain '/' which cannot be written as worksheet name
def pathNameCleanUp(path):
    head, tail = os.path.split(path)
    if tail != '':
        return tail
    else:
        os.path.basename(head)

class XlsUtil:
    # VV: a bit of a sketch, needs testing (test cases defined in
    def pathNameCleanUp(self, pathToFile):
        # strip out dir path and extension
        return os.path.splitext(os.path.basename(pathToFile))[0]

    def generateDeviationsSheetName(self, fP1, fP2):
        # we want the differences sheet name to incorporate release names so that it's possible to
        # combine deviation sheets from different reports into a single workbook
        # a bit of logic to strip out common prefix (if any), plus extension
        # doesn't deal with suffix atm
        fP1 = self.pathNameCleanUp(self, fP1)
        fP2 = self.pathNameCleanUp(self, fP2)
        cmnPref = os.path.commonprefix([fP1, fP2])
        if len(cmnPref) > 0:
            theMatch = re.compile('^[^0-9]*').match(cmnPref)
            # print('theMatch: {}'.format(theMatch))
            if theMatch:
                # print('cmnPref-before: {}'.format(cmnPref))
                cmnPref = cmnPref[:theMatch.end()]
                # print('cmnPref-after: {}'.format(cmnPref))
            fP1 = fP1[len(cmnPref):]
            fP2 = fP2[len(cmnPref):]
        return "{}-vs-{}".format(fP1, fP2)


# parsing argument
parser = argparse.ArgumentParser()
parser.add_argument('--curStat', help="path to current statistics report")
parser.add_argument('--prevStat', help="path to previous statistics report")
parser.add_argument('--outputFile', help="path to the output file")

# TODO: improve the args check so that:
# if number of arguments is not what is expected (3), print help message for users
if len(sys.argv) == 1:
    parser.print_help()
    sys.exit(1)

args = parser.parse_args()

w = Writer(pathNameCleanUp(args.outputFile))

reportCur = Report(args.curStat)
reportPrev = Report(args.prevStat)

# write deviation report to the 1st worksheet
w.writeDiffReport(reportCur, reportPrev)
w.writeReport(reportCur)
w.writeReport(reportPrev)

# close writer
w.close()
