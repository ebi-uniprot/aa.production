#!/usr/bin/env python

"""generateReleaseStatisticsDeviationReport.py: Generates 'deviation' report showing how much each AA Production Statistics number deviated from previous release."""
__author__ = "Cherry Hanquez"
__copyright__ = "Copyright 2018, The European Bioinformatics Institute (EMBL-EBI)"
__credits__ = ["Cherry Hanquez"]
__status__ = "Development"


# file arguments:
# --prevStat	path to previous statistics report
# --curStat	path to current statistics report
# --outputFile	path to the output file

# Requirements (can be done later on)
# the script must check that:
# * both input files: 
# ** a) exist;
# ** b) we have read access for them;
# ** c) non-empty (let's say, at least 100 bytes);
# * that output file does NOT exist;
# * that we can write into suggested output location (e.g. by trying to create an empty output file first).

