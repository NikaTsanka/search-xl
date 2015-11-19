import openpyxl


def get_files():
    wb = openpyxl.load_workbook('example.xlsx', read_only=True)
    print("Available Sheets")

    sheets = wb.get_sheet_names()
    print(sheets)

    sheet_name = input("Which sheet would you like to work with: ")

    if sheet_name in sheets:
        sheet = wb.get_sheet_by_name(sheet_name)
        for i in range(1, 8, 2):
            print(i, sheet.cell(row=i, column=2).value)


if __name__ == "__main__":
    get_files()
