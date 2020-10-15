#!/usr/bin/python3

# to do:
# * initial working script (✅ 2020-10-15)

# Parser for command-line options, arguments and sub-commands
import argparse

# check for boolean value
def str2bool(v):
    if isinstance(v, bool):
        return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')

# check command-line arguments
class Formatter(argparse.ArgumentDefaultsHelpFormatter, argparse.RawDescriptionHelpFormatter):
    pass
parser = argparse.ArgumentParser(
    description='Parse scanned Nexus PDF reports for certain tags and outputs an XLSX file.\nE.g.:\n  parser.py -p path/to/my/reports -t Tumorfläche Grading -o results.xlsx -v True',
    formatter_class=Formatter)
parser.add_argument(
    '-p', '--path',
    default='documents',
    help='set path to the folder that contains the scanned Nexus PDF reports')
parser.add_argument(
    '-t', '--tags',
    nargs='+',
    default=['Tumorfläche', 'Grading'],
    help='set the search tags')
parser.add_argument(
    '-o', '--output',
    default='results.xlsx',
    help='set output filename')
parser.add_argument(
    '-v', '--verbose',
    type=str2bool,
    nargs='?',
    const=True,
    default=False,
    help='more output while the script is running')
arguments = parser.parse_args()

print('Starting script:')

# import other modules
if arguments.verbose:
    print('  importing modules…', end='')

# Creating Excel XLSX files
import xlsxwriter

# Object-oriented filesystem paths
import pathlib

# System-specific parameters and functions
import sys

# Mathematical functions
import math

# Pythonic API for parsing PDF files
from pdfreader import PDFDocument, SimplePDFViewer

# Regular expression operations
import re

if arguments.verbose:
    print(' done')

# initital settings
if arguments.verbose:
    print('  initializing settings…', end='')

# initial settings
files = dict()
counter = dict()
counter['all'] = 0
counter['tags'] = 0
counter['nothing'] = 0
counter['incomplete'] = 0
counter['current'] = 0
tags = dict()
maxLengths = dict()
maxLengths['source'] = 0
for tag in arguments.tags:
    counter[tag] = 0
    maxLengths[tag] = 0

if arguments.verbose:
    print(' done')

# get file list
if arguments.verbose:
    print('  getting file list…', end='')

fileList = pathlib.Path(arguments.path).glob('**/*.pdf')

if arguments.verbose:
    print(' done')

if arguments.verbose:
    print('  parsing files:')
for item in fileList:
    print('    ' + str(item))
    files[str(item)] = dict()
    if len(str(item)) > maxLengths['source']:
        maxLengths['source'] = len(str(item))

    # open PDF report file
    fd = open(item, 'rb')
    viewer = SimplePDFViewer(fd)

    # iterating over all pages
    for pageNumber in range(1, viewer.doc.root['Pages']['Count'] + 1):
        viewer.navigate(pageNumber)
        viewer.render()

        # get all lines of the page
        allLines = viewer.canvas.text_content.split('\n')

        # filter for lines with 'readable' text content only
        cleanLines = []
        for line in allLines:
            lineMatch = re.match('\((.*) \)', line)
            if lineMatch is not None:
                cleanLines.append(lineMatch.group(1))
            lineMatch = re.match('\[\((.*) \) -\d+\.\d+ \((.*) \)\]', line)
            if lineMatch is not None:
                cleanLines.append(lineMatch.group(1))
                cleanLines.append(lineMatch.group(2))

        for tag in arguments.tags:
            if tag + ':' in cleanLines:
                files[str(item)][tag] = cleanLines[cleanLines.index(tag + ':') + 1]
                if len(str(cleanLines[cleanLines.index(tag + ':') + 1])) > maxLengths[tag]:
                    maxLengths[tag] = len(str(cleanLines[cleanLines.index(tag + ':') + 1]))
                counter[tag] += 1
                counter['tags'] += 1
                counter['current'] += 1
            elif not tag in files[str(item)].keys():
                files[str(item)][tag] = 'not found'
                if len(str(files[str(item)][tag])) > maxLengths[tag]:
                    maxLengths[tag] = len(str(files[str(item)][tag]))
    if counter['current'] == 0:
        counter['nothing'] += 1
        counter['incomplete'] += 1
    elif counter['current'] < len(arguments.tags):
        counter['incomplete'] += 1
    counter['current'] = 0
    counter['all'] += 1

# exit if no files were found
if counter['tags'] == 0:
    print('[WARNING] No PDF files found in folder \'' + arguments.path + '\'')
    sys.exit('Exiting')
if arguments.verbose:
    print('  total files parsed: ' + str(counter['all']))
    print('  total tags found: ' + str(counter['tags']))
    for tag in arguments.tags:
        print('    ' + tag + ': ' + str(counter[tag]))
    print('  total files with missing tags found: ' + str(counter['incomplete']))
    print('  total files without any tags found: ' + str(counter['nothing']))

# create an XLSX workbook
if arguments.verbose:
    print('  creating XLSX workbook…', end='')
workbook = xlsxwriter.Workbook(arguments.output)
if arguments.verbose:
    print(' done')

# table headers
workbookHeader = ['#', 'source']
for tag in arguments.tags:
    workbookHeader += [tag]

# start in second row
row = 1
col = 0

# add worksheet
worksheet = workbook.add_worksheet()

# declare different formats
headerBold = workbook.add_format({'bold': True})
headerBoldRight = workbook.add_format({'bold': True, 'align': 'right'})
numberPercent = workbook.add_format({'num_format': '##0 %'})
numberSpace = workbook.add_format({'num_format': '### ### ##0'})
for i in range(len(workbookHeader)):
    worksheet.write(0, i, workbookHeader[i], (headerBoldRight if (i == 0 or i == 2) else headerBold))

# freeze first row
worksheet.freeze_panes(1, 0)

# set autofilter
worksheet.autofilter(0, 0, 1, len(workbookHeader) - 1)

# write rows
if arguments.verbose:
    print('  writing rows…', end='')
for item in files:
    worksheet.write(row, col, row, numberSpace)
    worksheet.write(row, col + 1, item)
    index = 2
    for keys in files[item]:
        if re.match('\d+', files[item][keys]):
            worksheet.write(row, col + index, int(files[item][keys]) / 100, numberPercent)
        else:
            worksheet.write(row, col + index, str(files[item][keys]))
        index += 1
    row += 1
if row == 1:
    worksheet.write(row, col, 'no files found')

# adjust column widths
worksheet.set_column(0, 0, len(str(row)) + math.floor((len(str(row)) - 1) / 3) + 3)
worksheet.set_column(1, 1, max(maxLengths['source'], len(workbookHeader[1])) + 2)
for i in range(0, len(arguments.tags)):
    worksheet.set_column(i + 2, i + 2, max(maxLengths[arguments.tags[i]], len(workbookHeader[i + 2])) + 2)

# close workbook
workbook.close()
if arguments.verbose:
    print(' done')

print('All done and exiting.\n')
