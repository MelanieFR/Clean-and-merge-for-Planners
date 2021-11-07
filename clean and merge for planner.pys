import pandas as pd
from datetime import date

file_mps = 'C:\\Users\\ms3504\\Desktop\\Python Script\\Planner_file\\PLN_WW_O.xls.xlsx'
file_seg = 'C:\\Users\\ms3504\\Desktop\\Python Script\\Planner_file\\211011 - ABC RST Segmentation.xlsx'

mps = pd.read_excel(file_mps)
seg = pd.read_excel(file_seg)

print(mps.head())

# Trim the space trailed in the columns of product file, Ex: 'Item           ' => 'Item'
mps.columns = mps.columns.str.rstrip()

# Trim the space trailed in all rows of Item column, Ex: '8085-00481     ' => '8085-00481'
mps['Item'] = mps['Item'].str.rstrip()

# Print the first 10 rows the product panda object
print(mps.head(10))

# Create the column "Value" on the MPS file to see the final ordering value
mps['Value'] = mps['Frz Cost']*mps['Opn Qty']

# Trim the space trailed in the columns of inventory file, Ex: 'Item CD ' => 'Item CD'
seg.columns = seg.columns.str.rstrip()
print(seg.head())

# Trim the space trailed in all rows of all columns, Ex: '8085-00481 ' => '8085-00481', 'A1X  ' => 'A1X'
seg['Item CD'] = seg['Item CD'].str.strip()
seg['Segm'] = seg['Segm'].str.rstrip()
seg['RST'] = seg['RST'].str.rstrip()

# Check for duplicate values in seg file. 
print(seg.duplicated().value_counts())

# Merge seg file into the mps based on the 'Item' column in mps file and 'Item CD' column in seg.
# This is equivalent to the vlookup function of Excel.
merged_file = mps.merge(seg, left_on = 'Item', right_on='Item CD', how='left')
print(merged_file.head(20))

# Check the type of each column
print(merged_file.info())

# Only select the columns you are interested in.
merged_file = merged_file[['Pref', 'Item', 'Rel Dte', 'Due Dte', 'Res Dte', 'Opn Qty', 'Lot Qty', 'Lead Time',
                  'Item Type', 'Vendr','Segm', 'RST', 'Description', 'Bu', 'Div', 'Mrkt',
                  'Frz Cost', 'Value', 'Fac', 'Pln', 'On Hand Qty', 'Opn MPS', 'R/O Trans', 'Opn CST', 'Avg Sales', 'Mth Cov','P/O Transit',
                           'U/M Conv']]

print(merged_file.columns)

# Because Pandas will append the time into the date field, it will become something like 27/10/2021 12:00:00 AM. I only want to keep the date, no time.
merged_file['Rel Dte'].dt.date
merged_file['Due Dte'].dt.date
print(merged_file.head())

# Sort the merged file by Rel Dte 
merged_file.sort_values(['Rel Dte'], inplace=True)

print(merged_file.head())

# remove the rows with Vendr = 0
merged_file = merged_file[merged_file.Vendr != 0]
print(merged_file.head())

# Save the merged_file to an excel file
# Rename the file name before to run the code
merged_file.to_excel('C:\\Users\\ms3504\\Desktop\\Python Script\\Planner_file\\mps_ready_051121.xlsx', index=False)
