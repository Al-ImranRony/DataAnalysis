import openpyxl
import pandas
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

wb = openpyxl.load_workbook("KPMG_dataset_sprocket_central.xlsx")
#TODO: Get the sheets name of the excel file. 
Sheets = []
Sheets = wb.get_sheet_names()
i = 0
for sheet in Sheets:
    i += 1
    print(str(i)+'.', sheet, end='  ')

#TODO: Find the number of records & columns in a sheet
Sheet4 = wb.get_sheet_by_name("CustomerDemographic")
print(Sheet4.max_row)
dateCount = 0                                       
for row in Sheet4.rows:                             # Finding date data in the column
    if row[5] is not None:
        dateCount += 1
print(dateCount)

Sheet2 = wb.get_sheet_by_name("Transactions")
dateCount = 0
for row in Sheet2.rows:
    if row[3] is not None:
        dateCount += 1
print(dateCount)

Sheet5 = wb.get_sheet_by_name("CustomerAddress")
print(Sheet5.max_row)

#TODO: Checking if there has an empty cell or not
emptyCnt = 0
for rowNum in range(2, Sheet4.max_row):
    if Sheet4.cell(row=rowNum, column=3).value is None:
        emptyCnt+=1
if emptyCnt > 0:
        print(emptyCnt, "Empty cells in column", 3)