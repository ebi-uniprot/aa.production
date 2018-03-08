#!/usr/bin/env python

"""generateReleaseStatisticsDeviationReport.py: Generates 'deviation' report showing how much each AA Production Statistics number deviated from previous release."""
__author__ = "Cherry Hanquez"
__copyright__ = "Copyright 2018, The European Bioinformatics Institute (EMBL-EBI)"
__credits__ = ["Cherry Hanquez"]
__status__ = "Development"

import sys
import argparse

from readData import Report
from writeData import Writer

# parsing argument
parser = argparse.ArgumentParser()
parser.add_argument('--curStat', help="path to current statistics report")
parser.add_argument('--prevStat', help="path to previous statistics report")
parser.add_argument('--outputFile', help="path to the output file")

# improve the args check so that:
# if number of arguments is not what is expected (3), print help message for users
# if len(sys.argv) == 1 or len(sys.argv) != 7

args = parser.parse_args()
def error_args(s):
    print(s)
    parser.print_help()
    sys.exit(1)

if args.outputFile is None:
    error_args("Please type in the output file after \' --outputFile \'")

if args.curStat is None:
    error_args("Please input the current statistics report after \' --curStat \'")

if args.prevStat is None:
    error_args("Please input the previous statistics report after \' --prevStat \'")


#print(len(sys.argv))
#print(str(sys.argv))

#w = Writer(pathNameCleanUp(args.outputFile))

reportCur = Report(args.curStat)
reportPrev = Report(args.prevStat)

w = Writer(args.outputFile)
# write deviation report to the 1st worksheet
w.writeDiffReport(reportCur, reportPrev)
w.writeReport(reportCur)
w.writeReport(reportPrev)

# close writer
w.close()
