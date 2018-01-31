#!/usr/bin/env python

"""generateReleaseStatisticsDeviationReport.py: Generates 'deviation' report showing how much each AA Production Statistics number deviated from previous release."""
__author__ = "Cherry Hanquez"
__copyright__ = "Copyright 2018, The European Bioinformatics Institute (EMBL-EBI)"
__credits__ = ["Cherry Hanquez"]
__status__ = "Development"

import sys
import xlsxwriter
import argparse

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

class Worksheet:
	def __init__(self, workbook, name):
		self.worksheet = workbook.add_worksheet(name)
		self.row = 0

	def append(self, s):
		# write headers
		self.worksheet.write(self.row, 0, s.name)
		for col in range(0, len(s.headers)):
			self.worksheet.write(self.row, col + 1, s.headers[col])
		# from the next row, write the data
		self.row += 1
		col = 0
		for n in (convertList(s.data)):
			for col in range(0, len(n)):
				self.worksheet.write(self.row, col, n[col])
			self.row += 1
		self.row += 1


class Writer:
	def __init__(self, filename):
		self.workbook = xlsxwriter.Workbook(filename)

	# take a report class and write to a worksheet
	def writeReport(self, report):
		worksheet = Worksheet(self.workbook, report.name)
		for s in report.listOfSections:
			worksheet.append(s)

	def close(self):
		self.workbook.close()

class Report:
	def __init__(self, filename):
		self.name = filename
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
				print("writing section " + s.name)
				self.listOfSections.append(s)

			in_file.close()


parser = argparse.ArgumentParser()
parser.add_argument('--curStat', help="path to current statistics report")
parser.add_argument('--outputFile', help="path to the output file")
parser.add_argument('--prevStat', help="path to previous statistics report")

args = parser.parse_args()

w = Writer(args.outputFile)

r1 = Report(args.curStat)
w.writeReport(r1)
print('********')
r2 = Report(args.prevStat)
w.writeReport(r2)

# close writer
w.close()
