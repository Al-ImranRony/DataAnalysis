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

#TODO: Find the number of records 7 columns of the sheet
print(sheet0.shape)
print(sheet3.shape)
print(sheet4.shape)
print(sheet1.shape)
print(customer_data.shape)

# print(sheet3.head())

#TODO: Distinct records
customerIds = sheet1.iloc[:,1].unique()
distIds = sheet1.iloc[:,1].nunique()    # First value of customerIds will be column name,
print(customerIds[0], distIds-1)        # so minus it from distIds 
                                        




