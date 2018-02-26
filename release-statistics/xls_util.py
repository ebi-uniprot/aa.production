import os.path
import re


class XlsUtil:
    # VV: a bit of a sketch, needs testing (test cases defined in
    def pathNameCleanUp(self, pathToFile):
        # strip out dir path and extension
        return os.path.splitext(os.path.basename(pathToFile))[0]
        # Clean-up of path name so it doesn't contain '/' which cannot be written as worksheet name
        # head, tail = os.path.split(pathToFile)
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
        cmn_pref = os.path.commonprefix([fp1, fp2])
        if len(cmn_pref) > 0:
            theMatch = re.compile('^[^0-9]*').match(cmn_pref)
            # print('theMatch: {}'.format(theMatch))
            if theMatch:
                # print('cmnPref-before: {}'.format(cmnPref))
                cmn_pref = cmn_pref[:theMatch.end()]
                # print('cmnPref-after: {}'.format(cmnPref))
            fp1 = fp1[len(cmn_pref):]
            fp2 = fp2[len(cmn_pref):]
        return "{}-vs-{}".format(fp1, fp2)
