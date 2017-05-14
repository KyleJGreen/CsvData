import ParserFunctions

def main():
    reportName = "CensusCommunityDataReport"    # name of the file for generating the report
    directories = [r"/home/jupiter/Work/Community Mapping/regionTree/"]  # directories from which we are pulling the files
    csvFiles = []   # list of all csv files
    csvPaths = []   # list of all csv files and their path
    fieldsList = []  # list of all fields from csv file
    csvDict = {}    # dictionary of all csv files and their fields

    # fill lists for both the csv files and their paths to analyze for the report
    csvFiles, csvPaths = ParserFunctions.fillCsvLists(directories)
    # fill dictionary with the name of the csv file as the key and a list of the fields it contains as its value
    csvDict = ParserFunctions.fillCsvDict(csvFiles, csvPaths)
    # generate non-duplicate list of all fields from the csv files
    fieldsList = ParserFunctions.mergeDictLists(csvDict)
    # generate a report of all files per field and all fields per file
    ParserFunctions.generateReport(fieldsList, csvDict, reportName, directories[0])

if __name__ == '__main__': main()