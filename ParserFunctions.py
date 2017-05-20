from os import listdir
from os.path import isfile, join
import xlrd
import csv
import xlsxwriter

## generates a list of all files from a given directory
from xlrd import XLRDError

# retuns a list of files given a directory
def pullFiles(path):
    return returnCsvs([f for f in listdir(path) if isfile(join(path, f))], path)

# takes a list of file names and returns only those ending in a .txt file format, .csv file format, or .xls,
# which is converted to a .csv using the excelToCsv function
def returnCsvs(files, path):
    csvFiles = []
    # for every file in our list, check the extension, and add it to our CSV Files list if it is a .txt
    for file in files:
        isExtension = 0  # boolean value for determining whether or not we are at the beginning of the file extension
        extension = ""  # the name of the file's extension
        fileName = ""
        for index, char in enumerate(file):
            # if we are at the extension part of the filename, append characters to the extension variable
            if isExtension is 1:
                extension = extension + char
            # otherwise, write to the filename
            else:
                if char != ".":
                    fileName = fileName + char
            # set isExtension to true once the '.' char is reached and we are not at the beginning of the file name
            if char is '.' and index is not 0:
                isExtension = 1
        # convert .xls files to .csv format
        if extension == "xls" or extension == "xlsx":
            # see if file is compatible with the excelToCsv function
            try:
                excelToCsv(file, fileName + ".csv", path)
                csvFiles.append(fileName + ".csv")
            except XLRDError:
                print(file + " failed to properly convert to .csv format")  # sometimes this error prints and file still converts

        # add .txt files to the CSV Files list
        if extension == "txt" or extension == "csv":
            csvFiles.append(file)
    return csvFiles

# takes a list of directories and adds all csv files from directories into two lists
def fillCsvLists(directories):
    csvFiles = []   # list of all .csv files without their path
    csvPaths = []   # list of all .csv files with their paths

    # iterate over all directories and add their files to the lists of csv files and paths
    for directory in directories:
        newFiles = pullFiles(directory)  # create list of files for each directory
        # add all csv files from each directory to the csv files list
        for file in newFiles:
            csvPaths.append(directory + file)
            csvFiles.append(file)

    return csvFiles, csvPaths

# fill dictionary with the name of the csv file as the key and a list of the fields it contains as its value
def fillCsvDict(csvFiles, csvPaths):
    csvDict = {}

    # iterate through both dictionaries to fill the new dictionary with the file name as the key and a list of its fields as the value
    for file, path in zip(csvFiles, csvPaths):
        #  open the current file
        with open(path) as f:
            lines = f.readlines()   # read the lines from the file
            # make sure that lines isn't empty
            if lines:
                # for .csv files converted from .xls format
                if "Table with row headers" not in lines[0]:
                    csvDict[file] = parseLine(lines[0])  # parse line from .csv file
                else:
                    csvDict[file] = parseLine(lines[2])  # parse line from .csv file
    f.close()   # close the file
    return csvDict

# parses a line from a csv file into an array of strings
def parseLine(line):
    row = []    # a list for all fields contained in a line of a .csv file
    field = ""  # field to be assigned to the list

    # iterate over all characters in a given line, parsing on ',' and adding to the row list
    for char in line:
        # skip over quotation marks
        if char == "\"":
            continue
        if char is not "," and char != '\n':
            field = field + char    # append char to the field String
        else:
            row.append(field.lower())   # append field to the row list
            field = ""  # reset field
    return row

# generate non-duplicate list of all fields from the csv files
def mergeDictLists(csvDict):
    newList = []    # new non-duplicate list

    # merge all fields into a non-duplicate list
    for key, value in csvDict.iteritems():
        newList = mergeLists(value, newList)

    return newList

# merges two lists, eliminating duplicates
def mergeLists(listA, listB):
    myset = set(listA).union(set(listB))
    return sorted(list(myset))

# for converting .xls and .xlsx file to .csv format
def excelToCsv(excelFile, csvFile, path):
    xls = xlrd.open_workbook(path + excelFile, on_demand=True)  # pointer to the .xls(x) file
    sheetNames = xls.sheet_names()  # sls(x) sheet name list
    csvF = open(path + csvFile, 'wb')   # pointer to a new writable .csv file
    wr = csv.writer(csvF, quoting=csv.QUOTE_ALL)    # .csv file writer

    # iteratate over all sheets in sheet names list, adding the, to a workbook and a worksheet
    for sheet in sheetNames:
        workbook = xlrd.open_workbook(path + excelFile)
        worksheet = workbook.sheet_by_name(sheet)

        # write rows to .csv file
        for rownum in xrange(worksheet.nrows):
            wr.writerow(
                list(x.encode('utf-8') if type(x) == type(u'') else x
                    for x in worksheet.row_values(rownum)))

    csvF.close()    # close the .csv file

# generate a .xls report containing which fields are contained in which files and which files are contained in which fields
def generateReport (fieldList, csvDict, reportName, path):
    reportDict = {}  # dictionary for generating the report
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(path + reportName + ".xlsx")
    worksheet = workbook.add_worksheet()
    row = 2  # Set initial row index to 2 for writing the .xlsx file
    # Add a bold, underline, and bold/underline format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    underline = workbook.add_format({'underline': True})
    boldUnderline = workbook.add_format({'underline': True, 'bold': True})

    # add all fields from the field list as keys to our new dictionary, with an empty list as their value
    for field in fieldList:
        reportDict[field] = []

        # fill the lists in the report dictionary with all files that contain the given field
        for file, fields in csvDict.iteritems():
            for element in fields:
                if field == element:
                    reportDict[field].append(file)

    # Write header for FIELD --> FILES Table of the worksheet
    worksheet.write(0, 0, "FIELD --> FILES", boldUnderline)
    worksheet.write(1, 0, "FIELD", bold)
    worksheet.write(1, 1, "FILES", bold)

    # Fill the FIELD --> FILES Table of the .xls(x) file
    for field, files in sorted(reportDict.iteritems()):
        col = 0
        worksheet.write(row, col, field, bold)
        col += 1

        for file in files:
            worksheet.write(row, col, file)
            col += 1
        row += 1

    row += 1

    # Write header for FILE --> FIELDS Table of the worksheet
    worksheet.write(row, 0, "FILE --> FIELDS", boldUnderline)
    row += 1
    worksheet.write(row, 0, "FILE", bold)
    worksheet.write(row, 1, "FIELDS", bold)
    row += 1

    # Fill the FILE --> FIELDS Table of the .xls(x) file
    for file, fields in sorted(csvDict.iteritems()):
        col = 0
        worksheet.write(row, col, file, bold)
        col += 1

        for field in fields:
            worksheet.write(row, col, field)
            col += 1
        row += 1

    workbook.close()
