# Classes for reading in data and separate them into sections
from xls_util import XlsUtil

# Report class for opening a file and return a list of sections
class Report:
    def __init__(self, path):
        self.name = XlsUtil().pathNameCleanUp(path)
        self.listOfSections = []
        with open(path, 'rt') as in_file:
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
                numbers.append(float(i))
        self.data.append((" ".join(x[:data_start]), numbers))

# separate the file into sections whereas an empty line
def parseSection(in_file):
    dataLines = list()
    while 1:
        line = in_file.readline().rstrip('\n')
        if not line or line == "\n":
            break
        dataLines.append(line)

    if len(dataLines) == 0:
        return None

    s = Section(dataLines[0])
    for l in dataLines[1:]:
        s.append(l)
    return s
