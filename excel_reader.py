# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 11:41:13 2018

@author: NIRAV.KAKKAD
"""

import os
import xlrd
import sys


class Excel_reader():
    rowno, colno = 0, 0

    def __init__(self):
        print("Reading Excel File")

    def open_excel(self, filename):
        workbook = xlrd.open_workbook(filename)
        sheet = workbook.sheet_by_index(0)
        return sheet

    def create_text_from_excel(self, textfilename, excelobj, rowscount, colscount):
        print("hello")
        self.i = 0
        with open(textfilename, 'w') as f1:
            while (self.i < rowscount):
                print(5)
                self.j = 0
                while (self.j < colscount):
                    f1.write(str(excelobj.cell(self.i, self.j).value))
                    f1.write("\t")
                    self.j += 1
                self.i += 1
                f1.write("\n")


if __name__ == "__main__":

    excelfile = sys.argv[1]
    textfilename = sys.argv[2]
    obj = Excel_reader()
    excelsheet = obj.open_excel(excelfile)
    print("Number of cols", excelsheet.ncols)
    print("Number of rows", excelsheet.nrows)

    i, j = 0, 0
    while (i < excelsheet.nrows):
        j = 0
        while (j < excelsheet.ncols):
            print(excelsheet.cell(i, j).value)
            j += 1
        i += 1
    print("hi")
    obj.create_text_from_excel(textfilename, excelsheet, excelsheet.nrows, excelsheet.ncols)
