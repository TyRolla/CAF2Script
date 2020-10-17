from openpyxl import load_workbook
from openpyxl import Workbook

#Refer to 1.1
outputWb = Workbook()
wb1 = load_workbook('fullListHTC.xlsx')
wb2 = load_workbook('CAF2Aparcels.xlsx')

#Refer to 1.2
filepath = "./output.xlsx"
sheetNew = outputWb.active
sheetOne = wb1['Hawaii Addresses']
sheetTwo = wb2['Hawaii']

#Refer to 1.3
maxRow1 = sheetOne.max_row
maxRow2 = sheetTwo.max_row
output = []


def iter_rows(ws, row):
    for row in ws.iter_rows():
        yield [cell.value for cell in row]


for i in range(2, maxRow1):
    currID1 = sheetOne.cell(row=i, column=1)
    print(currID1.value)
    for j in range(2, maxRow2):
        currID2 = sheetTwo.cell(row=j, column=2)
        if int(currID2.value) > int(currID1.value):
            break
        if int(currID1.value) == int(currID2.value):
            print("Match " + str(currID1.value))
            output.append(list(sheetOne[i]))

interArr = []
returnArr = []
for row in range(0, len(output)):
    for cell in range(0, len(output[row])):
        interArr.append(output[row][cell].value)
        if cell == len(output[row]) - 1:
            returnArr.append(tuple(interArr))
            interArr = []

sheetNew.append(('Parcel Number', 'Type', 'Address', 'County', 'GeoID', 'Latitude', 'Longitude', 'CAF2A_block_id',
                 'CAF2A_status', 'stateabbr'))
for row in returnArr:
    sheetNew.append(row)
outputWb.save(filepath)

print("SEARCH COMPLETE! PLEASE CHECK FOLDER WITH SCRIPT FOR OUTPUT.XLSX")
