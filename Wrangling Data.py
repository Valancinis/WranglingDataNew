from datetime import timedelta
import pandas as pd
import openpyxl

# _Functions_

def saveExcel():
    navdf.to_excel('output excel.xlsx')
    print('Excel file has been successfully saved')

# _End of Functions_

# read files
importdataOOR = pd.read_excel(r'OOR for Thermo Fisher Scientific - LT.xlsx')
importdataOPO = pd.read_excel(r'LED OOR Query.xlsx')

# picking needed data
oordf = importdataOOR.iloc[:, [2, 10, 16]].copy()
navdf = importdataOPO.iloc[:, [0, 2, 3]].copy()

oordf.rename({"Material": "ID", "Current ship date": "Delivery date"}, axis='columns', inplace=True)

# setting 'Delivery date' column datatype to datetime
oordf['Delivery date'] = pd.to_datetime(oordf['Delivery date'], errors='coerce', infer_datetime_format=True)
navdf['Expected Receipt Date'] = pd.to_datetime(navdf['Expected Receipt Date'], errors='coerce', infer_datetime_format=True)

# add transit time to suppliers delivery dates (+8 days in LED case)
for i, row in oordf.iterrows():
    oordf.at[i, 'Delivery date'] = (oordf.at[i, 'Delivery date'] + timedelta(days=8))

navdf['Ddelta'] = (navdf['Expected Receipt Date'] - oordf['Delivery date']) / timedelta(days=1)

# Add a new column
navdf['Update'] = ''
navdf = navdf.merge(oordf, how = 'inner', on = ['PO','ID'])

# looping through the dataframe, turning timedelta objects to int and updating "Update" column values if row needs to be updated in the ERP system
for i in navdf.index:
    n = navdf.loc[i, 'Ddelta']
    if (n > 9) or (n < -9):
        navdf.loc[i, 'Update'] = "True"
    else:
        navdf.loc[i, 'Update'] = "False"

#print(mergeddf.head())

# _Exporting files
#saveExcel()
