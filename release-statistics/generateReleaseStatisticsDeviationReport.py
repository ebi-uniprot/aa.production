#!/usr/bin/env python

"""generateReleaseStatisticsDeviationReport.py: Generates 'deviation' report showing how much each AA Production Statistics number deviated from previous release."""
__author__ = "Cherry Hanquez"
__copyright__ = "Copyright 2018, The European Bioinformatics Institute (EMBL-EBI)"
__credits__ = ["Cherry Hanquez"]
__status__ = "Development"

import sys
import xlsxwriter
import argparse

#f1 = input ("Enter the first file name: ")
#f2 = input ("Enter the second file name: ")

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
                numbers.append(i)
        self.data.append((x[:data_start], numbers))

# separate the file into sections whereas an empty line
def parseSection(s):
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
def convertList(listOfData):
    l = []
    for (name, numbers) in listOfData:
        if len(name) > 1:
            n = " ".join(name)
        else:
            n = name[0]
        numbers.insert(0, n)
        l.append(numbers)
    return l

class Writer:
    def __init__(self, filename):
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet()
        self.row = 0

    def append(self, s):
		# write headers
        self.worksheet.write(self.row, 0, s.name)
        for col in range(0, len(s.headers)-1):
            self.worksheet.write(self.row, col + 1, s.headers[col])
        # from the next row, write the data
        self.row += 1
        col = 0
        for n in (convertList(s.data)):
            for col in range(0, len(n)):
                self.worksheet.write(self.row, col, n[col])
            self.row += 1
        self.row += 1

    def close(self):
        self.workbook.close()


parser = argparse.ArgumentParser()
parser.add_argument('--curStat', help="path to current statistics report")
parser.add_argument('--outputFile', help="path to the output file")
parser.add_argument('--prevStat', help="path to previous statistics report")

args = parser.parse_args()

print(args.curStat)

with open(args.curStat, 'rt') as in_file:
    # skip first 6 lines
    for i in range(1,6):
        in_file.readline()

    # read data and write to excel
    w = Writer(args.outputFile)
    while True:
        s = parseSection(in_file)
        if s is None:
            break
        print("writing section " + s.name)
        w.append(s)

    # close file descriptor and writer
    w.close()
    in_file.close()
