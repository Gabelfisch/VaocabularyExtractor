# Documentation:
# https://pypdf2.readthedocs.io/en/3.x/

import pdfplumber
import xlsxwriter as XLWriter
from random import randint

filepath = "AllVocables.PDF" # Insert file path of the PDF with the vocabulary
sides = (1, 3) # Insert from which page to which page you want to export the Vobaulary
               # Warning!!! Is not alway taking exactly the end of the side, can't fix it though bc its a bug in xslxwriter.


vocabTable = []
isThreeTimesPossible = False
isFourTimesPossible = False


# Opens the File (PDF) with the Vocabulary
with pdfplumber.open(filepath) as temp:

    # adding all the (selected) sides to one big chart (vocabTable)
    for i in range(sides[0]-1, sides[1]):
        vocabTable.extend([[f"----- page {i+1} ----- " , "", "---------  ---------", ""]])
        vocabTable.extend(temp.pages[i].extract_table())


# Creating the Excel File
workbook = XLWriter.Workbook('VocabularyList.xlsx')
worksheet = workbook.add_worksheet()


# Overwriting the not needed nulls in the chart with ''
vocabularyTableCopy = vocabTable

for j, table in enumerate(vocabularyTableCopy):
    for i, field in enumerate(table):
        if (field == None):
            vocabTable[j][i] = ''


j = 0 # Row/Index of the Table in Excel
i = 0 # Actuel inxex of the for loop/vocabTable
for definition, rnd, vocab, rn2 in vocabTable:

    if (definition != '' or  vocab != ''): 
        if (definition == ''): 
            j -= 1
            
            if (isThreeTimesPossible):
                worksheet.write(j, 1, f"{vocabTable[i-2][2]} {vocabTable[i-1][2]} {vocab}")
                isFourTimesPossible = True

            elif (isFourTimesPossible):
                worksheet.write(j, 1, f"{vocabTable[i-3][2]} {vocabTable[i-2][2]} {vocabTable[i-1][2]} {vocab}")
            

            else:
                worksheet.write(j, 1, f"{vocabTable[i-1][2]} {vocab}")
                isThreeTimesPossible = True
        else:
            worksheet.write(j, 0, definition)
            worksheet.write(j, 1, vocab)
            isThreeTimesPossible = False
            isFourTimesPossible = False
        j+=1
    i+=1


# Closes and saves the Excel file
workbook.close()

print("finished")