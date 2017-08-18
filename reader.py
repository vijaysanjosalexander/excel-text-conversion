# -*- coding: utf-8 -*-
"""
Written by: Vijay Sanjos Alexander
Email: vijaysanjosalexander@gmail.com
Copyright (c) 2017, VSA.
License: MIT (see LICENSE for details)
"""

import xlrd
import datetime
import os
import time
import codecs
start_time = time.time()
# Directory for files
file_dir = "D:/Users/Folder/"


# Function that convert cell value to string
def convtostr(val, bool):
    if type(val) is unicode:
        return val.encode('utf-8')
    elif type(val) is float and bool == 0:
        return str(int(val))
    elif type(val) is float and bool == 1:
        xldate = xldate_as_datetime(val, 0)
        return xldate


# Function that convert float value to string
def xldate_as_datetime(xldate, datemode):
    # datemode: 0 for 1900-based, 1 for 1904-based
    dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=xldate + 1462 * datemode)
    return dt.strftime('%m/%d/%Y')

# Get all excel files
for f in os.listdir(file_dir):
    if f.endswith(".xlsx"):
        path = os.path.join(file_dir, f)
        print "Parsing file: {}".format(path)
        # Open the work book
        book = xlrd.open_workbook(path)
        # Get sheet count in excel
        tabCount = book.nsheets
        nlp_text = ""
        # Get file name without extension
        ff = os.path.splitext(f)[0]
        # Output file generated
        file_name = "{0}nlp_file-{1}.txt".format(file_dir,ff)
        print "Writing to file: {}".format(file_name)
        nlp_file = codecs.open(file_name, 'w', encoding='utf-8')
        # Iterate through worksheets and get values
        for i in range(0, tabCount):
            sheet = book.sheet_by_index(i)
            nlp_text += "SHEET NAME is {} and ".format(sheet.name)
            # Get the header values
            header = sheet.row(0)
            for rx in range(1, sheet.nrows):
                for cx in range(0, sheet.ncols):
                    headerVal = convtostr(header[cx].value, 0)
                    if 'date' in headerVal.lower():
                        cellVal = convtostr(sheet.cell_value(rx, cx), 1)
                    else:
                        cellVal = convtostr(sheet.cell_value(rx, cx), 0)
                    if cellVal !="":
                        nlp_text += "{0} is {1} and ".format(headerVal, cellVal)
                        nlp = nlp_text.replace('\n', ' ')
                nlp += "\n"
        nlp_file.write(nlp.decode('utf-8'))
        nlp_file.close()
etime = time.time() - start_time
print("--- %s seconds ---" % etime)


def secondsToStr(t):
    return "%d:%02d:%02d.%03d" % reduce(lambda ll, b: divmod(ll[0], b) + ll[1:], [(t*1000,), 1000, 60, 60])
print(secondsToStr(etime))

