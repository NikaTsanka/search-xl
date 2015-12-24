## Search Excel
Search Excel is a small Python script that can search structured Excel spreadsheets. It uses Python third-party module “Openpyxl”

## Supported Platforms:
- Windows 7 or later versions.
- Most of the linux distributions.
- OS X 10.9 "Mavericks" or later versions.

## Requirements:
- Python 3.4 or higher on all platforms
- Openpyxl 2.3.1

## How to Install:
The primary way to install third-party modules is to use Python’s pip tool.

- The executable file for the pip tool is called pip on Windows and pip3 on OS X and Linux. On Windows, you can find pip at C:\Python34\Scripts\pip.exe. On OS X, it is in /Library/Frameworks/Python.framework/Versions/3.4/bin/pip3. On Linux, it is in /usr/bin/pip3. While pip comes automatically installed with Python 3.4 on Windows and OS X, you must install it separately on Linux.

- Install openpyxl using pip. It is advisable to do this in a Python virtualenv without system packages:

```$ pip install openpyxl```

## How to Run:
On all systems input files can be passed at the command line or at the runtime.
- Windows: `python search_xl.py InputFile.xlsx`
- Linux / OS X: `python3 search_xl.py InputFile.xlsx`

## How it works:
- Search Excel uses a third-party library called Openpyxl. Openpyxl is a Python library for reading and writing Excel 2010 xlsx/xlsm/xltx/xltm files.

- First program opens user specified workbook and informs the user about available spreadsheets inside the workbook. Next it will ask the user to enter a search value and name of the sheet to lookup. The search value is only checked for specified column search range which can be chnaged. After finding the search value results are save in $HOME/Desktop directory in text format. User can keep looking for the same value in different spreadsheets available in the workbook or can terminate the program. Upon termination final sum of taxa abundance from all sheets will be written back on the file.

## How to Edit:
- To edit the search column range simply change the constant values in code:

```
Line 12: START_COL = 1
Line 13: END_COL = 7
```
- To include the last two columns(SUM/AVE) change code on line 69:

`max_col = sheet.max_column   -  1`

to

`max_col = sheet.max_column  + 1`

## Input Files:
Input excel files must meet the following criteria:
- “A” column must contain “#” sign. “#” is used to determine where the starting point of the search rectangle is in the spreadsheet.

## Output:
Output file structure:

|Sheet Name:   |   |   |   |
|---|---|---|---|
|   |Location 1:   |Location 2:   |Location 3:   |
|Bacteria Name: #/Coordinates    |Value   |Value   |Value   |
|Separator:    |_   |_   |_   |
|Sum of Each Column:   |Location 1:   |Location 2:   |Location 3:   |
|   |Sum 1   |Sum 2   |Sum 3   |
|Sheet Overall:   |Total Sum   |   |   |

|Bacteria Name: Taxa Abundance: Sum   |
|---|

## Author:
+ Nika Tsankashvili: [Github](https://github.com/NikaTsanka) , [Twitter](https://twitter.com/NikaTsanka)

## Contrubutors:
+ Jessica Joyner: [Github](https://github.com/jjoyner07)
