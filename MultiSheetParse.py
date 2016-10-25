#!/usr/bin/python

import pandas as pd


headers = ["SheetName"]


sourceFile = "SourceFile.xlsx"


def getStats(excelFile):
    wholeBook = pd.ExcelFile(sourceFile)
    sheetNames = wholeBook.sheet_names

    with open('SheetNames.txt', 'w') as w:
        for i in range(0, len(sheetNames)):
            w.writelines(str(sheetNames[i]) + ", ")

    w.close()    
    print(sheetNames)
    print(len(sheetNames))

    nRecords = 0
    for i in range(0, len(sheetNames)):
        if "Employee-" in str(sheetNames[i]):
            nRecords+=1
    print("The number of records found is: ", nRecords)



def parseSheet(nsheets):
    for i in range(0, nsheets):
        # get just the files we needs
        sheet1 = wholeBook.parse(i)
        colnames = sheet1.columns

        rows = {}

        for j in range(0, len(sheet1.index)):
            rows[j] = []
            for b in range(0, len(colnames)):
                rows[j].append(sheet1[colnames[b]][j])

            # output to DF, check data types
            # test first, if wrong then fail & halt

        # get all headers (robust against length variation in row 2)
        headers = []
        for j in range(0, len(colnames)):
            # output[str(colnames[i])] = rows[0][i]
            headers.append(colnames[j])
        for j in range(0, len(rows[1])):
            # output[str(row[1][i])] 
            headers.append(rows[1][j])




    # functioning example load into dataframe
    d = {}
    for i in range(0, len(headers)):
        print(headers[i])
        if i < len(colnames):
            
            d[headers[i]] = rows[0][i] #make this a series!
        if i >= len(colnames):
            seriesVals = []
            for b in range(2, len(rows)):
                seriesVals.append(rows[b][i-len(colnames)])
            d[headers[i]] = seriesVals

    newFrame = pd.DataFrame(d)        
    print(newFrame)







def parseAll(sourceFile):

    # read in the excel file
    wholeBook = pd.ExcelFile(sourceFile)

    # run loop for all sheets
    # for i in range(0, len(validSheets)):
    #   

    # stats on the file size
    sheetNames = wholeBook.sheet_names
    nsheets = (len(sheetNames))
    print("Total number of sheets: ", nsheets)







    # write to file
    newFrame.to_excel('output.xlsx', "It Worked!")