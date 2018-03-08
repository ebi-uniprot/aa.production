# compare the data from two different reports

from readData import Report, Section

# def is_num_list(l):
#     if len(l) == 0:
#         return
#     for i in l:
#         if not (type(i) == type(int()) or type(i) == type(float())):
#             raise ValueError

class DiffReport:
    def __init__(self, report1, report2):
        self.diffSec = []

        for sec1 in report1.sections:
            for sec2 in report2.sections:
                if sec1.name == sec2.name:
                    self.diffSec.append((sec1.name, sec1.headers, DiffSection(sec1.data, sec2.data)))

class DiffSection:
    def __init__(self, d1, d2):
        self.diffSec = []

        # i = 0
        # while i < len(d1):
        #     (lineName1, nb1) = d1[i]
        #     (lineName2, nb2) = d2[i]
        #     is_num_list(nb1)
        #     is_num_list(nb2)
        #     if lineName1 == lineName2:
        #         self.diffSec.append((lineName1, nb1, nb2))
        #    else:
        #         self.diffSec.append(d1[i])
        #         self.diffSec.append(d2[i])
        #    i += 1

        # TODO can also try doing a dictionary and then put it back to diffSec
        threeWayDiff = []
        for (linename1, nb1) in d1:
            cp = [nb1, None]
            threeWayDiff.append((linename1, cp))

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

        for (l, v) in threeWayDiff:
            if v[0] == None:
                self.diffSec.append((l, v[1]))
            elif v[1] == None:
                self.diffSec.append((l, v[0]))
            else:
                self.diffSec.append((l, v[0], v[1]))
        #
        # for (linename1, nb1) in d1:
        #     for(linename2, nb2) in d2:
        #         if linename2 == linename1:
        #             is_num_list(nb1)
        #             is_num_list(nb2)
        #             self.diffSec.append((linename1, nb1, nb2))
        #         #else:
        #         #    self.diffSec.append((linename1, nb1))
        #         #    self.diffSec.append((linename2, nb2))
        #
        # for(linename2, nb2) in d2:
        #     if linename2 in self.diffSec:
        #         pass
        #     else:
        #         self.diffSec.append((linename2, nb2))
