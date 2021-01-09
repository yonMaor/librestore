#!/usr/bin/python3

import csv
from openpyxl import load_workbook
import sys
import warnings

excel_extensions = [".xlsx", '.xlsm', '.xlsb', '.xltx', '.xltm', '.xls', '.xlt', '.xls']

def csv_from_excel(book, sheet):
    workbook = load_workbook(filename=book, read_only=True)
    sheet = workbook[sheet]
    csv_file = open('e.csv', 'w')
    record = csv.writer(csv_file, quoting=csv.QUOTE_ALL)

    for row in sheet.rows:
        index = []
        for cell in row:
            index.append(cell.value)
        record.writerow(index)

    csv_file.close()

#Check number of arguments
if len(sys.argv) > 3:
    warnings.warn("Too many arguments have been sent to convert.py. Using the first two arguments.")

#Check file extension (first argument)
if len(sys.argv[1]) > 4:
    if ((sys.argv[1][-5:]) in excel_extensions) or ((sys.argv[1][-4:]) in excel_extensions):
        pass
    else:
        warnings.warn("The excel file does not have a standard excel extension.")
else:
    warnings.warn("The excel file does not have a standard excel extension.")

csv_from_excel(sys.argv[1], sys.argv[2])
