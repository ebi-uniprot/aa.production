# compare the data from two different reports

from readData import Report, Section

class DiffReport:
    def __init__(self, report1, report2):
        self.diffSec = []

        for sec1 in report1.sections:
            for sec2 in report2.sections:
                if sec1.name == sec2.name:
                    self.diffSec.append((sec1.name, sec1.headers, DiffSection(sec1.data, sec2.data)))
                if sec1.is_footer == True or sec2.is_footer == True:
                    break
class DiffSection:
    def __init__(self, d1, d2):
        self.diffSec = []

        # can also try doing a dictionary and then put it back to diffSec
        threeWayDiff = []

        # copy data1 in a list first
        for (linename1, nb1) in d1:
            cp = [nb1, None]
            threeWayDiff.append((linename1, cp))

        # look for the same line name in data2
        for(linename2, nb2) in d2:
            found = None
            for (l, v) in threeWayDiff:
                if l == linename2:
                    found = v
                    break

            if found == None:
                threeWayDiff.append((linename2, [None, nb2]))
            else:
                found[1] = nb2

        # copy back data back to diffSec, depending on if there is any new data line in data2
        for (l, v) in threeWayDiff:
            if v[0] == None:
                self.diffSec.append((l, v[1]))
            elif v[1] == None:
                self.diffSec.append((l, v[0]))
            else:
                self.diffSec.append((l, v[0], v[1]))
