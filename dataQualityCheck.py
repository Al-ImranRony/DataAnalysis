import sys
import pandas as pd


# sys.stdin = open("output.txt", "w")

excel_file = 'KPMG_dataset_sprocket_central.xlsx'
customer_data = pd.read_excel(excel_file)

print(customer_data.shape)
