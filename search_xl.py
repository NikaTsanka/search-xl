#!/usr/bin/python3
import sys
import openpyxl
from openpyxl.utils import get_column_letter

__author__ = 'Nika Tsanakshvili'
__email__ = "nikatsanka@gmail.com"
__version__ = "1.1"
__copyright__ = "Copyright (C) 2015 Nika Tsankashvili"
__license__ = "Public Domain"

# Search Range
START_COL = 1
END_COL = 7


def search_xl():
    # open the workbook
    # Before assignments
    global start_row, wb

    # if length of the argv is more than 1 then the argument is passed.
    if len(sys.argv) > 1:
        # path should be at 1 because .py is at 0
        workbook = sys.argv[1]
    else:
        workbook = input("Please enter name of the workbook: ")

    input_val = True
    while input_val:
        try:
            wb = openpyxl.load_workbook(workbook, read_only=True)
            input_val = False
        except FileNotFoundError:
            print("Specified workbook cannot be located.")
            workbook = input("Please enter name of the workbook: ")

    print("Available Sheets")
    # print sheets
    sheets = wb.get_sheet_names()
    print(sheets)
    # get the search value.
    search_val = input("What are you looking for? ")
    # ask usr for more sheets.
    user_ = True
    # keep checking sheets
    while user_:
        # get working sheet.
        sheet_by_name = input("Which sheet would you like to work with: ")
        # if chosen sheet is in the workbook then do the following
        if sheet_by_name in sheets:
            # set the sheet
            sheet = wb.get_sheet_by_name(sheet_by_name)

            # get maximum number of rows in a sheet.
            max_row = sheet.max_row

            # find the starting point
            # start looking for # sign in the first column
            for k in sheet.columns[0]:
                # start checking
                cell_val = k.value
                if cell_val == '#':
                    start_row = k.row
                    break
            # save number of results
            count_res = 0

            # start looking for search value
            # 1 to 7 columns that's first 1
            for i in range(START_COL, END_COL):
                # start_row to max_row. STEP 1
                for j in range(start_row, max_row, 1):
                    # check for search value
                    if search_val == sheet.cell(row=j, column=i).value:
                        print("[ Found at row:", j, "column:", get_column_letter(i), "]")
                        #print(sheet.cell(row=j, column=i).coordinate)
                        count_res += 1

            print("Total number of \'", search_val, "\' found: ", count_res, sep='')
            ans = input("Would you like to search another sheet? (Y/n): ")
            if ans == 'N' or ans == 'n':
                user_ = False
            else:
                print("Available Sheets")
                # print sheets
                sheets = wb.get_sheet_names()
                print(sheets)
                # else it will ask the user to enter the sheet name again.
        else:
            print("\'", sheet_by_name, "\' is not in the workbook.", sep='')


if __name__ == "__main__":
    if sys.version_info > (3, 4):
        search_xl()
    else:
        print("This software requires Python 3.4 or higher to run." +
              "\nYou can download the latest version of Python from: https://www.python.org/")
