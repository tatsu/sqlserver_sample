#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import sys
from enum import Enum

import xlrd

import sqlserver as server


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
    db = server.SqlServer()
    db.open(server.DRIVER, server.SERVER, server.DATABASE, server.USERNAME,
            server.PASSWORD)

    try:
        sheet = excel.sheet_by_name('Sheet1')
        found = False
        for y in range(0, sheet.nrows):
            row = sheet.row(y)
            if not found:
                if row[Column.ID.value].value == 'id':
                    found = True
            else:
                id = int(row[Column.ID.value].value)
                tup = (id,
                       row[Column.NAME.value].value,
                       row[Column.FILENAME.value].value,)
                if db.is_record_existed(id):
                    writer.writerow(('REJECTED', *tup))
                else:
                    db.insert_record(*tup)
                    writer.writerow(('INSERTED', *tup))

    except xlrd.XLRDError as e:
        print(e)
        raise

    db.close()
    file.close()


if __name__ == '__main__':
    args = sys.argv
    if len(args) != 2:
        print('Usage:\n    python insert_excel_record.py filename.xlsx')
        exit(1)

    do_excel(args[1])
    print('Finish')
