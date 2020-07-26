import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime


excel_file = "KPMG_datasets.xlsx"

sheet0 = pd.read_excel(excel_file, 'Title Sheet', skiprows=1)
sheet3 = pd.read_excel(excel_file, sheet_name='CustomerDemographic', skiprows=1)
sheet4 = pd.read_excel(excel_file, sheet_name='CustomerAddress', skiprows=1)
sheet1 = pd.read_excel(excel_file, sheet_name='Transactions', skiprows=1)

customer_data = pd.concat([sheet3, sheet4, sheet1])
# If excelFile has lot of sheets !
# sheets = []
# for sheet in datasets.sheet_names:
#    sheets.append(datasets.parse(sheet))
# customer_data = pd.concat(sheets)

#TODO: Find the number of records 7 columns of the sheet
print(sheet3.shape)
print(sheet4.shape)
print(sheet1.shape)
print(customer_data.shape)

# print(sheet3.head())

#TODO: Distinct records
customerIds = sheet1.iloc[:,1].unique()
distIds = sheet1.iloc[:,1].nunique()    # First value of customerIds will be column name,
print(customerIds[0], distIds-1)        # so minus it from distIds 

#TODO: Convert character-date to Date
charDate = "July 25 2020 12:01 AM"
dateObj = datetime.strptime(charDate, "%B %d %Y %H:%M %p")
print(dateObj)

#TODO: Data exploration
srtV = sheet3.sort_values(['tenure', 'DOB'], ascending=False).head(1000)
print(srtV)

#TODO: Plot the output 
srtV.plot()
plt.show()

# Histogram plot
fig = plt.figure()
ax = fig.add_subplot(1,1,1)

ax.hist(srtV['DOB'],bins=5)
plt.title('Age observation')
plt.xlabel('DOB')
plt.ylabel('#Numbers')
plt.show()

# Scatter plot
fig = plt.figure()
ax = fig.add_subplot(1,1,1)

ax.scatter(sheet3['tenure'],sheet3['customer_id'])
plt.title('Tenure observation')
plt.xlabel('tenure')
plt.ylabel('#customer_id')
plt.show()