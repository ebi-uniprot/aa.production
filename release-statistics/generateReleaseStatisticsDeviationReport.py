#!/usr/bin/env python

"""generateReleaseStatisticsDeviationReport.py: Generates 'deviation' report showing how much each AA Production Statistics number deviated from previous release."""
__author__ = "Cherry Hanquez"
__copyright__ = "Copyright 2018, The European Bioinformatics Institute (EMBL-EBI)"
__credits__ = ["Cherry Hanquez"]
__status__ = "Development"

import sys
import xlsxwriter
import argparse
import os.path

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
		self.data.append((" ".join(x[:data_start]), numbers))

# Worksheet class for writing individual reports and differetial report
class Worksheet:
	def __init__(self, workbook, name):
		self.worksheet = workbook.add_worksheet(name)
		self.row = 0

	def print_headers(self, name, headers):
		# write headers
		self.worksheet.write(self.row, 0, name)
		for col in range(0, len(headers)):
			self.worksheet.write(self.row, col + 1, headers[col])


	def append(self, s):
		self.print_headers(s.name, s.headers)
		# from the next row, write the data
		self.row += 1
		col = 0
		for n in convertList(s.data):
			for col in range(0, len(n)):
				self.worksheet.write(self.row, col, n[col])
			self.row += 1
		self.row += 1

	def write_list(self, col, l):
		for c in range(0, len(l)):
			self.worksheet.write(self.row, col + c, l[c])
		return (col + len(l))

	def appendDiff(self, diffSec):
		(name, headers, diffData) = diffSec
		self.print_headers(name, headers * 4)
		self.row += 1
		for line in diffData.diffSec:
			col = 0
			# when there is a difference in name, only write one set of data
			if len(line) == 2:
				(lineName, nb) = line
				self.worksheet.write(self.row, col, lineName)
				col += 1

			# write two sets of data with the same name
			elif len(line) == 3:
				(lineName, nb1, nb2) = line
				self.worksheet.write(self.row, col, lineName)
				col += 1
				col = self.write_list(col, nb1)
				col = self.write_list(col, nb2)
				v = []
				p = []
				for i in range(0, len(nb1)):
					diffVal = int(nb1[i]) - int(nb2[i])
					v.append(diffVal)
					if int(nb2[i]) == 0:
						p.append(0.0)
					else:
						diffPer = "{:.2%}".format(diffVal / int(nb2[i]))
						p.append(diffPer)
				col = self.write_list(col, v)
				col = self.write_list(col, p)
			else:
				print("error")
			#for col in range(0, len(line)):
			#	self.worksheet.write(self.row, col, line[col])
			self.row += 1
		self.row += 1

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
		worksheet = Worksheet(self.workbook, "compReport")
		diffR = DiffReport(r1, r2)

		for diffSec in diffR.diffSec:
			worksheet.appendDiff(diffSec)

	def close(self):
		self.workbook.close()

# Report class for opening a file and return a list of sections
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
				#print("writing section " + s.name)
				self.listOfSections.append(s)

			in_file.close()


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
			if lineName1 == lineName2:
				self.diffSec.append((lineName1, nb1, nb2))
			else:
				self.diffSec.append(d1[i])
				self.diffSec.append(d2[i])
			i += 1

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
		#numbers.insert(0, name)
		l.append([name] + numbers)
	return l

# Clean-up of path name so it doesn't contain '/' which cannot be written as worksheet name
def pathNameCleanUp(path):
	head, tail = os.path.split(path)
	if tail != '':
		return tail
	else:
		os.path.basename(head)

# parsing argument
parser = argparse.ArgumentParser()
parser.add_argument('--curStat', help="path to current statistics report")
parser.add_argument('--prevStat', help="path to previous statistics report")
parser.add_argument('--outputFile', help="path to the output file")

# If there is no arguments keyed in, print help message for users
if len(sys.argv) == 1:
	parser.print_help()
	sys.exit(1)

args = parser.parse_args()

w = Writer(pathNameCleanUp(args.outputFile))

r1 = Report(pathNameCleanUp(args.curStat))
w.writeReport(r1)
print('********')
r2 = Report(pathNameCleanUp(args.prevStat))
w.writeReport(r2)

# write differetial report to one of the worksheets
w.writeDiffReport(r1, r2)

# close writer
w.close()
