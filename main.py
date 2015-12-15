"""
Open and read the cells of an Excel document with the openpyxl module.
Calculate all the tract and population data and store it in a data structure.
Write the data structure to a text file with the .py extension using the pprint module.
"""

import openpyxl, pprint
from openpyxl.utils import get_column_letter


def get_files():
    # open the workbook
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

    search_val = input("So what are you looking for? ")

    sheet_by_name = input("Which sheet would you like to work with: ")
    # if chosen sheet is in the workbook then do the following
    if sheet_by_name in sheets:
        # set the sheet
        sheet = wb.get_sheet_by_name(sheet_by_name)
        # this will print the first row.
        for i in range(1, 7):
            for j in range(1, 8, 1):
                # print(sheet.cell(row=i, column=j).value)
                cell_val = sheet.cell(row=j, column=i).value
                if search_val == cell_val:
                    print("Category: ", sheet.cell(row=1, column=i).value)
                    print("[ Found at row:", j, "col:", i, "]")
                    # print entire row
                    col = sheet.max_column + 1
                    for rVAl in range(j, col):
                        print("Name: ", sheet.cell(row=1, column=rVAl).value)
                        print(str(sheet.cell(row=j, column=rVAl).value))

    # # set the sheet again
    # sheet = wb.get_sheet_by_name(sheet_by_name)
    #
    # # get total number of rows and columns
    # rows = sheet.max_row
    # cols = sheet.max_column
    #
    # # get max letters
    # # max_let_row = get_column_letter(sheet.max_row)
    # max_let_col = get_column_letter(sheet.max_column)
    #
    # # print results
    # print("\nIn sheet: ", sheet_by_name, " There are: ", rows,
    #       " rows and ", cols, " cols", sep='')
    #
    # # convert to strings for args
    # start = 'A1'  # constant
    # end = str(max_let_col + str(rows))
    #
    # # print values/entire table
    # for rowOfCellObjects in sheet[start:end]:
    #     for cellObj in rowOfCellObjects:
    #         print(cellObj.coordinate, cellObj.value)
    #     print('----End of row----')

if __name__ == "__main__":
    get_files()
