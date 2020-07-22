import sys
import pandas as pd


excel_file = "KPMG_dataset_sprocket_central.xlsx"
datasets = pd.ExcelFile(excel_file)

sheet0 = pd.read_excel(datasets, 'Title Sheet')
sheet3 = pd.read_excel(excel_file, sheet_name=3, index_col=0)
sheet4 = pd.read_excel(excel_file, sheet_name=4, index_col=0)
sheet1 = pd.read_excel(excel_file, sheet_name=1, index_col=0)

customer_data = pd.concat([sheet3, sheet4, sheet1])
# If excelFile has lot of sheets !
# sheets = []
# for sheet in datasets.sheet_names:
#    sheets.append(datasets.parse(sheet))
# customer_data = pd.concat(sheets)

print(sheet0.shape)
print(sheet3.shape)
print(sheet4.shape)
print(sheet1.shape)
print(customer_data.shape)

print(sheet3.head())

#TODO: Checking if there has an empty cell or not
for rowNum in range(2, sheet3.max_row):
    print(sheet3.cell(row=rowNum, column=2).value)
        # if cell is None: print("Empty cell.")