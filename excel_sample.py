#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import sys
from enum import Enum

import xlrd


class Column(Enum):
    ID = 1
    NAME = 2
    FILENAME = 3


def do_excel(filename):
    try:
        excel = xlrd.open_workbook(filename)
    except IOError as e:
        print(e)
        raise

    file = open('result.csv', 'a')
    writer = csv.writer(file)

    try:
        sheet = excel.sheet_by_name('Sheet1')
        found = False
        for y in range(0, sheet.nrows):
            row = sheet.row(y)
            if not found:
                if row[Column.ID.value].value == 'id':
                    print('{:>3} {:20} {:30}'.format(
                        row[Column.ID.value].value.upper(),
                        row[Column.NAME.value].value.upper(),
                        row[Column.FILENAME.value].value.upper(),
                    ))
                    found = True
            else:
                tup = (int(row[Column.ID.value].value),
                       row[Column.NAME.value].value,
                       row[Column.FILENAME.value].value,)
                print('{:3d} {:20} {:30}'.format(*tup))
                writer.writerow(('INSERTED', *tup))
    except xlrd.XLRDError as e:
        print(e)
        raise

    file.close()


if __name__ == '__main__':
    args = sys.argv
    if len(args) != 2:
        print('Usage:\n    python excel_sample.py filename.xlsx')
        exit(1)

    do_excel(args[1])
    print('Finish')
