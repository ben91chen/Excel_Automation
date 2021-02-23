import pandas as pd
import os
import shutil
import datetime

#get files
path = os.getcwd()
files = os.listdir(path)
spreadsheets = [f for f in files if f[-3:] == 'csv']
print(spreadsheets)

# create a workbook
date = datetime.datetime.now().strftime("%Y%m%d")
writer = pd.ExcelWriter('{}NetsuiteReports.xlsx'.format(date))

for spreadsheet in spreadsheets:
    # Create pandas dataframe from spreadsheet csv
    df = pd.read_csv(spreadsheet)
   # df = df.style.set_properties(**{'font-size': '12pt'}) Having issues with this line
    print(df)

    # Add sheet to workbook
    sheet = spreadsheet.split('Results')[0]
    df.to_excel(writer, sheet_name=sheet, index=False)
    print(f"{spreadsheet} was written to sheet {sheet}.")

    # adjust column widths
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheet].set_column(col_idx, col_idx, column_length)

# Save workbook
writer.save()

#move XLSX file to target folder, and CSVs to Archive folder
source = '/Users/benchen/Desktop/Netsuite-Weekly-Reports'
dest1 = '/Users/benchen/Desktop/Netsuite-Weekly-Reports/Target-Folder'
dest2 = '/Users/benchen/Desktop/Netsuite-Weekly-Reports/Archive'

final_files = os.listdir(source)
for x in final_files:
    #if x[-4:] == 'xlsx':
    #    shutil.move(x, dest1)
    if x[-3:] == 'csv':
        shutil.move(x, dest2)



