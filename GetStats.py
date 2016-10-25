#!/usr/bin/python

import pandas as pd

# file pathway. eventually change to a sys input
sourcePath = 'SourceFile.xlsx'

def makeDF(sheet1, colnames, index):
    rows = {}
    for i in range(len(sheet1.index)):
        rows[i] = []
        for b in range(len(colnames)):
            rows[i].append(sheet1[colnames[b]][i])

    # get all headers (robust against variation in row 2)
    headers = []
    for i in range(len(colnames)):
        # output[str(colnames[i])] = rows[0][i]
        headers.append(colnames[i])
    for i in range(len(rows[1])):
        # output[str(row[1][i])] 
        headers.append(rows[1][i])

    # functioning example load into dataframe
    d = {"SheetName":str(sheetNames[index])}
    for i in range(len(headers)):
        print(headers[i])
        if i < len(colnames):
            
            d[headers[i]] = rows[0][i] #make this a series!
        if i >= len(colnames):
            seriesVals = []
            for b in range(2, len(rows)):
                seriesVals.append(rows[b][i-len(colnames)])
            d[headers[i]] = seriesVals

    newFrame = pd.DataFrame(d)
    return newFrame


# read in file
wholeBook = pd.ExcelFile(sourcePath)
# get all names of sheets, and write to reference file
sheetNames = wholeBook.sheet_names
with open('SheetNames.txt', 'w') as w:
    for i in range(len(sheetNames)):
        w.writelines(str(sheetNames[i]) + ", ")

w.close()    

nRecords = 0
nExtra = 0
nTotal = len(sheetNames)
frames = []
goalSheet = input("What tab do you want to parse on?")


# for all sheets:
for i in range(0, len(sheetNames)):
    # if the sheet is an employee record:
    if goalSheet in str(sheetNames[i]):
        nRecords += 1
        # add dictionary
        
        # open this sheet, grab top column names
        thisSheet = wholeBook.parse(i)
        thisColNames = thisSheet.columns

        # make a dataframe out of this sheet
        # add the dataframe to the list
        frames.append(makeDF(thisSheet, thisColNames, i))
        
    else:
        nExtra += 1

result = pd.concat(frames)
# print(result.columns)
proceed = input("regular or fancy?")
if proceed == "r":
    result.to_excel('outputFile.xlsx', "It Worked!")
if proceed == "f":
    writer = pd.ExcelWriter('outputFile.xlsx', engine='xlsxwriter')
    result.to_excel(writer, index=False, sheet_name="It's Fancy!")
    workbook = writer.book
    worksheet = writer.sheets["It's Fancy!"]
    total_fmt = workbook.add_format({'align': 'left', 'num_format': '##0',
                                 'bold': True, 'bottom':6})
    workbook.close()


print('The total number of sheets is: ', len(sheetNames))
print("The number of records found is: ", nRecords)
print("The number of non-record sheets is: ", nExtra)
print("The number of unaccounded sheets is: ", nTotal-nExtra - nRecords)
print("-------------------------------------------------")
print(result.columns)

# if __name__ == '__main__':
#     getFileStats(str(sys.argv))