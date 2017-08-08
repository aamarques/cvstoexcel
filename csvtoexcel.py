#!/usr/bin/python -W ignore
# -*- coding: UTF-8 -*-
#
# Read a csv and writes a excel xlsx file merging columns with equal values.
# This do not merge rows!!
# This version only process columns from A to Z, But can be easily modified.

# Feel free to adapt for you needs.
#
# Antonio Marques - July 2017 - aamarques@gmail.com

import sys
import csv
import xlsxwriter
import string
import argparse
from platform import python_version


# FUNCTIONS

def open_file(filename):
    try:
        filehandle = open(filename)
        return filehandle
    except IOError as e:
        print "I/O error({0}): {1} - {2}".format(e.errno, filename, e.strerror)
        exit()
    except:
        print "Unexpected error:", sys.exc_info()[0]
        exit()


def check_duplicated():
# Read te fistr line with field names
    fnames = csv.DictReader(csvfile, delimiter=_delimiter, dialect="excel").fieldnames
# control variables
    theend = len(fnames) - 1
    found_error = 0

# First loop in list elements
    for first_x in range(len(fnames)):
        if first_x == theend:
            break
# Second loop trhu actual position to position + 1cp cs 
        second_x = first_x +1
        while second_x < len(fnames):

# Found an error ?
            if fnames[first_x] == fnames[second_x]:
                print "Field Name duplicated : {} = {}".format(fnames[first_x], fnames[second_x])
                found_error += 1

            second_x += 1

    if found_error != 0:
        print " ( {} ) erros were found! Exiting... \n".format(found_error)
        csvfile.close()
        sys.exit()

    csvfile.seek(0)

# PARSER ARGS
# Parse of arguments
parser = argparse.ArgumentParser(
    description=" ",
    epilog="Ex: csvtoexcell.py -i inputfile -c 6 -d ';' "
    )
#
parser.add_argument("-i", help="The input file to be parsed. REQUIRED", action="store", dest="ifilename")
parser.add_argument("-o", help="The output file without extension. Default is outfile.xlsx", action="store", dest="ofilename", default="outfile")
parser.add_argument("-d", help="Delimite used in file. Default is space", action="store", dest="delimiter", default=" ")
parser.add_argument("-c", help="Number of last column to merge. Default is 4", action="store", dest="last_col", default=4, type=int)
parser.add_argument("-v", "-V", "--version", action="version", version="%(prog)s 1.7.05.17")

args = parser.parse_args()

_ifilename = args.ifilename
_ofilename = str(args.ofilename) + ".xlsx"
_delimiter = args.delimiter
_last_col = args.last_col
_ver = python_version()

# VARS

# Colors
_RED_font  = "\033[1;31;48m"
_NORM_bg   = "\033[0;37;48m"

first_row = 2   # Start from A2. A1 is the line head
last_row = 2
index = 0
col_field_aux = None
celpos = list(string.ascii_uppercase)

# MAIN

# Verify Python Version
if _ver < "2.7.0":
	print _RED_font + "You need python in 2.7.0 or superior." + _NORM_bg
	print "Your python version is = {0}".format(_ver)
	sys.exit()


# Call help if there is no filename
#if _name_file or _path_name_file is None:
if _ifilename is None:
    parser.parse_args(['-h'])
    sys.exit()

# Open CSV File
csvfile = open_file(_ifilename)
check_duplicated()

# CONVERT TO EXCEL FORMAT

# Creates a xlsx file
wb = xlsxwriter.Workbook(_ofilename)
# Creates a Sheet
ws = wb.add_worksheet('Cluster')
# Format to single cels
cell_format = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
# Format to merged cells
merge_format = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

# LOOP READ AND WRITE

# Read the CSV file and write it in excel format

reader = csv.reader(csvfile, delimiter=_delimiter)
for row_index, row in enumerate(reader):
    for cell_index, cell in enumerate(row):
        ws.write(row_index, cell_index, cell, cell_format)
        ws.set_column(row_index, cell_index, 13)


# LOOP TO MERGE COLUMNS

# Return to the begining of CSV file to loop for columns and mege it.
csvfile.seek(0)
#
reader = csv.DictReader(csvfile, delimiter=_delimiter, dialect="excel")
heads = reader.fieldnames

for header_field in heads:

    for col_field in reader:
#        print col_field
        if col_field_aux is None:
            col_field_aux = col_field[header_field]
            continue

        if col_field[header_field] == col_field_aux:
            last_row += 1
        else:
            cells = '{}{}:{}{}'.format(celpos[index], first_row, celpos[index], last_row)
#            print cells, col_field_aux
            last_row += 1
            first_row = last_row
            ws.merge_range(cells, col_field_aux, merge_format)
            col_field_aux = col_field[header_field]

    if last_row >= 4:
        cells = '{}{}:{}{}'.format(celpos[index], first_row, celpos[index], last_row)
        ws.merge_range(cells, col_field_aux, merge_format)

    csvfile.seek(0)  # return to begin and start merge (if need) the next column of csv file
    csvfile.next()   # read the headers
    first_row = 2    # Initialize Variables
    last_row = 2
    col_field_aux = None
    index += 1       # Jumps to next alphabetic letter

    # Reach to the default column or passed by args?
    if index == _last_col:
        break

# Close everything
csvfile.close()
wb.close()