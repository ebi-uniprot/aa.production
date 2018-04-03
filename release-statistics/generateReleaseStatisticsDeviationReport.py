#!/usr/bin/env python

"""generateReleaseStatisticsDeviationReport.py: Generates 'deviation' report showing how much each AA Production Statistics number deviated from previous release."""
__author__ = "Cherry Hanquez"
__copyright__ = "Copyright 2018, The European Bioinformatics Institute (EMBL-EBI)"
__credits__ = ["Cherry Hanquez"]
__status__ = "Development"

import argparse

from readData import Report
from writeData import Writer

# parsing argument
parser = argparse.ArgumentParser(description="AA Production Statistics Deviation Report")
parser.add_argument('prev_stat', help="path to previous statistics report")
parser.add_argument('cur_stat', help="path to current statistics report")
parser.add_argument('output_file', help="path to the output file")

# improve the args check so that:
# if number of arguments is not what is expected (3), print help message for users
# if len(sys.argv) == 1 or len(sys.argv) != 7

args = parser.parse_args()
# def error_args(s):
#     print(s)
#     parser.print_help()
#     sys.exit(1)
#
# if args.outputFile is None:
#     error_args("Please type in the output file after \' --outputFile \'")
#
# if args.curStat is None:
#     error_args("Please input the current statistics report after \' --curStat \'")
#
# if args.prevStat is None:
#     error_args("Please input the previous statistics report after \' --prevStat \'")


report_prev = Report(args.prev_stat)
report_cur = Report(args.cur_stat)

w = Writer(args.outputFile)
# write deviation report to the 1st worksheet
w.writeDiffReport(report_cur, report_prev)
w.writeReport(report_cur)
w.writeReport(report_prev)

# close writer
w.close()
