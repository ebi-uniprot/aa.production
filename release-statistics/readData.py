# Classes for reading in data and separate them into sections
from xls_util import XlsUtil
from txt_util import has_text, skip_first_section


# Report class for opening a file and return a list of sections
class Report:
    def __init__(self, path):
        self.name = XlsUtil().pathNameCleanUp(path)
        self.sections = []
        self.footers = []
        with open(path, 'rt') as in_file:

            skip_first_section(in_file)

            # read data
            while True:
                s = parse_section(in_file)
                if s is None:
                    break
                #print("writing section " + s.name)
                self.sections.append(s)
                if s.is_footer:
                    self.footers.append(s)

# create a class for a typical section which contains header and data parts
class Section:
    def __init__(self, name):
        # remove leading and trailing whitespaces plus trailing colon:
        self.name = name.strip().rstrip(':')
        self.headers = []
        self.data = []
        self.is_footer = False

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
                    self.headers.append(i.rstrip(':'))

        # data is in a format of list of tuples, each of which contains two lists of strings
        numbers = []
        for i in x[data_start:]:
            if i.isdecimal():
                numbers.append(float(i))
        self.data.append((" ".join(x[:data_start]), numbers))

class Footer:
    def __init__(self, name):
        self.name = name.strip().rstrip(':')
        self.longHeader = None
        self.headers = []
        self.data = []
        self.is_footer = True
        self.tremblEntries = []

    def append(self, line):
        x = line.split()
        data_start = 0

        # when the name is only 'Global', take the next line as section name
        # percentages are formula
        if self.name == "Global":
            if self.longHeader == None:
                self.longHeader = line.strip().rstrip(':')

            numbers = []
            if x[1].isdecimal():
                numbers.append(float(x[1]))
                self.data.append((x[0], numbers))

            if x[0] == "TrEmbl":
                self.tremblEntries.append(float(x[2]))
                lineName = " ".join(x[:2]).strip().rstrip(":")
                self.data.append((lineName, self.tremblEntries))

        else:
            while not x[data_start].endswith(":") and not x[data_start].isdecimal():
                data_start += 1

            number = x[data_start + 1].rstrip(",")
            numbers = []
            if number.isdecimal():
                lineName = " ".join(x[:(data_start + 1)]).strip().rstrip(":")
                numbers.append(float(number))
                self.data.append((lineName, numbers))

            for i, v in enumerate(x):
                if v == "TrEmbl":
                    tremblStart = i
                    v2 = x[tremblStart + 1]
                    lineNameList = [v, v2]
                    lineName = " ".join(lineNameList).strip().rstrip(":")
                    if x[tremblStart + 2].isdecimal():
                        self.tremblEntries.append(float(x[tremblStart + 2]))
                        self.data.append((lineName, self.tremblEntries))

# separate the file into sections whereas an empty line
def parse_section(in_file):
    # optimise by avoiding iterating over the lines twice
    s = None

    while 1:
        line = in_file.readline()
        # break when it reaches the end of the file or finds an empty line
        if not line or line == "\n":
            break

        # create a section object after each empty line,
        # using the first line as the name of the section,
        # then append lines until the next empty one.
        if s is None:
            if line.startswith("Global"):
                s = Footer(line)
            else:
                s = Section(line)
        else:
            s.append(line)

    return s
