import os.path
import re


class XlsUtil:
    # VV: a bit of a sketch, needs testing (test cases defined in

    def __init__(self):
        self.regex_cell_ref = re.compile('^([A-Z])([0-9]+)$')

    def pathNameCleanUp(self, pathToFile):
        # strip out dir path and extension
        fileName = os.path.splitext(os.path.basename(pathToFile))[0]
        pref = re.compile('[^0-9]*').match(fileName)
        if pref:
            p = fileName[:pref.end()]
            return fileName[len(p):]

        # Clean-up of path name so it doesn't contain '/' which cannot be written as worksheet name
        # head, tail = os.path.split(pathToFile)  # different from using splitext()
        # if tail != '':
        #     return tail
        # else:
        #     os.path.basename(head)

    def generate_deviations_sheet_name(self, fp1, fp2):
        # we want the differences sheet name to incorporate release names so that it's possible to
        # combine deviation sheets from different reports into a single workbook
        # a bit of logic to strip out common prefix (if any), plus extension
        # doesn't deal with suffix atm
        fp1 = self.pathNameCleanUp(fp1)
        fp2 = self.pathNameCleanUp(fp2)
        # cmn_pref = os.path.commonprefix([fp1, fp2])
        # if len(cmn_pref) > 0:
        #     theMatch = re.compile('^[^0-9]*').match(cmn_pref)
        #     # print('theMatch: {}'.format(theMatch))
        #     if theMatch:
        #         # print('cmnPref-before: {}'.format(cmnPref))
        #         cmn_pref = cmn_pref[:theMatch.end()]
        #         # print('cmnPref-after: {}'.format(cmnPref))
        #     fp1 = fp1[len(cmn_pref):]
        #     fp2 = fp2[len(cmn_pref):]
        return "{}-vs-{}".format(fp1, fp2)

    def span_range(self, range_start, col_span=1, row_span=1):
        range_start_parts = self.regex_cell_ref.match(range_start)

        if not range_start_parts:
            raise Exception('could not parse range start: {}'.format(range_start))

        if col_span < 1:
            raise Exception('colspan must be >=1. got: {}'.format(col_span))
        if row_span < 1:
            raise Exception('rowspan must be >=1. got: {}'.format(row_span))

        col_name = range_start_parts.group(1)
        # do a bit of validation:
        if len(col_name) > 1:
            raise Exception('method can work with single-char column name currently')
        col_name = col_name.upper()
        if (ord(col_name) < ord('A')) or (ord(col_name) > ord('Z')):
            raise Exception('col name is expected to be within A-Z range. Got: {}'.format(col_name))

        new_col_name_ord = ord(col_name) + col_span - 1 # minus one since our column already spans "1 column"
        if (new_col_name_ord < ord('A')) or (new_col_name_ord > ord('Z')):
            raise Exception('col name {} spanned by {} would fall outside of A-Z range'.format(col_name, col_span))
        new_col_name = chr(new_col_name_ord)

        row_num = int(range_start_parts.group(2))
        new_row_num = row_num + row_span - 1 # minus one since our row already spans "1 row"

        cells_range = '{}:{}{}'.format(range_start, new_col_name, new_row_num)
        print('range start: {}, col_span: {}, row_span: {}-->{}'.format(range_start, col_span, row_span, cells_range))
        return cells_range
