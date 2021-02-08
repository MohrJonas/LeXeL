import xlrd
import os
import json
from pyexcel_ods import get_data
from sys import argv
from enum import Enum
from abc import ABC, abstractmethod

writeToFile = False
parser = None


class SSRead(ABC):
    @abstractmethod
    def getRows(self):
        pass

    @abstractmethod
    def getColumns(self):
        pass

    @abstractmethod
    def getCell(self, x, y):
        pass


class ExcelReader(SSRead):
    def __init__(self, path):
        self.file = xlrd.open_workbook(path)

    def getRows(self):
        return self.file.sheet_by_index(0).nrows

    def getColumns(self):
        return self.file.sheet_by_index(0).ncols

    def getCell(self, x, y):
        return self.file.sheet_by_index(0).cell_value(x, y)


class ODSReader(SSRead):
    def __init__(self, path):
        self.file = get_data(path)[list(get_data(path))[0]]

    def getRows(self):
        return len(self.file)

    def getColumns(self):
        maxLen = 0
        for row in self.file:
            if(len(row) > maxLen):
                maxLen = len(row)
        return maxLen

    def getCell(self, x, y):
        try:
            return self.file[x][y]
        except IndexError:
            return ""

def checkArgs():
    if(len(argv) == 1):
        print("Usage: lexel [excel-file]")
        print("Usage: lexel [excel-file] [output]")
        exit(-1)
    if(len(argv) != 2 and len(argv) != 3):
        raise Exception("Invalid args")
    if(not os.path.isfile(argv[1])):
        raise Exception("File doesn't exist")
    global writeToFile
    writeToFile = (len(argv) == 3)
    global parser
    if(argv[1].endswith(".xls")):
        parser = ExcelReader(argv[1])
    elif(argv[1].endswith(".ods")):
        parser = ODSReader(argv[1])
    else:
        raise Exception("Unable to find proper parser")


def newCell(content):
    if(content is None):
        return ""
    return str(content)


def newRow(cells):
    return "\t\t" + " & ".join(cells) + "\\\\\n" + "\t\t\\hline\n"


def newHeader():
    return "\\begin{adjustbox}{width=1.2\\textwidth,center=\\textwidth}\n" + '\t\\begin{tabular}{' + ("|c" * parser.getColumns() + "|") + '}\n\t\t\\hline\n'


def newFooter():
    return "\t\\end{tabular}\n\\end{adjustbox}\n"


def newTable(sheet):
    return

def parse():
    out = ""
    out += newHeader()
    for row in range(0, parser.getRows()):
        cells = []
        for cell in range(0, parser.getColumns()):
            cells.append(newCell(parser.getCell(row, cell)))
        row = newRow(cells)
        out += row
    out += newFooter()
    return out


if(__name__ == "__main__"):
    checkArgs()
    if(writeToFile):
        with open(argv[2], "w") as file:
            file.write(parse())
    else:
        print(parse())
