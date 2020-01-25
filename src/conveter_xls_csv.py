'''
Contributers: 
    1. Vrushabh
    2. Sonali 
Makeing sample code to conveter Xlsx/Xls file to CSV 
'''
import os
import sys
import glob
import xlrd
import unicodecsv
import openpyxl
import pandas as pd


def xls_to_csv_convertor(xls_filename):
    # Extract the filename along with path without extension.
    csv_filename = xls_filename.rsplit('/', 1)[-1].rsplit('.', 1)[0]

    try:
        # It will load the workbook.
        wb = xlrd.open_workbook(xls_filename)

        # Check the number of sheets in the workbook.
        sh = wb.nsheets
        print("Number of sheets {}".format(sh))
        shName = wb.sheet_names()

        # Loop through the all the sheets.
        for sheet_number in range(sh):
            try:
                # Open the sheet by index.
                wsh = wb.sheet_by_index(sheet_number)

                # Filename to generate the CSV file.
                fileName = "output-" + shName[sheet_number] + ".csv"

                # Open the csv file in binary write mode.
                fh = open(fileName, "wb")
                # Uses unicodecsv, so it will handle Unicode characters.
                csv_out = unicodecsv.writer(fh, encoding='utf-8')

                # Loop through the rows of the sheet and write to csv file.
                for row_number in range(wsh.nrows):
                    csv_out.writerow(wsh.row_values(row_number))

                # Close the csv file.
                fh.close()

                print("CSV file created successfully.")

            except Exception as _e:
                print("Error creating CSV file.")
                print(sys.exc_info())
    except Exception as _e:
        print("Error opening the file.")
        print(sys.exc_info())


def dump_csv(file, outdir):
    df1 = pd.ExcelFile(file)
    for sheet in df1.sheet_names:
        df2 = pd.read_excel(file, sheet)
        df2['Date Filled'] = df2['Date Filled'].apply(
            lambda x: x.strftime('%d-%m-%Y'))
        print(df2["Date Filled"])
        filename = os.path.join('output_' + sheet + '.' + 'csv')
        df2.to_csv(filename)


if __name__ == "__main__":
    dump_csv(sys.argv[1], sys.argv[1])
    # dump_csv('C:/Users/vrusdesh/Documents/workspace_testcase/app7/src/',
    # 'C:/Users/vrusdesh/Documents/workspace_testcase/app7/src/')
    # xls_to_csv_convertor(sys.argv[1])
