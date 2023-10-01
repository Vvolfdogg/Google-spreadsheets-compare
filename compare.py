import gspread
from oauth2client.service_account import ServiceAccountCredentials
from colorama import init, Fore
#init(convert=True) #uncomment this line if you are using powershell/cmd

scope = ["https://www.googleapis.com/auth/spreadsheets.readonly","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.readonly","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("cred.json", scope)
client = gspread.authorize(creds)
global_counter = 0 #counter for different cells in all sheets

def compare_sheets(sh1, sh2, sheetName, font):
    print(font + "\nComparing " + sheetName + " sheet...")

    minSize = min(len(sh1), len(sh2)) #cells from this rows will be compared, rest of the rows are missing from one file
    local_counter = 0 #counter for different cells in one sheet
    cells = [] #this array will be returned with all different cell coordinates

    row = 0
    while row < minSize: #looping through rows

        col = 0
        colLetter = ord("A")
        colFirstLetter = ord("@")

        size1Col = len(sh1[row])
        size2Col = len(sh2[row])
        sizeCol = min(size1Col, size2Col)

        while col < sizeCol: #looping through columns

            if sh1[row][col] != sh2[row][col]: #if different

                if colFirstLetter == ord('@'): #if cell has a single letter coordinate e.g. A51, not AB51
                    cellName = chr(colLetter) + str(row + 1)
                else:
                    cellName = chr(colFirstLetter) + chr(colLetter) + str(row + 1)

                print("Cell: ", cellName, end = '') #coordinates
                print(font + " File1: " + sh1[row][col] + " / File2: " + sh2[row][col]) #value in both sheets

                local_counter += 1
                global global_counter
                global_counter += 1
                cells.append(cellName)

            col += 1

            if colLetter == ord("Z"):
                colLetter = ord("A")
                colFirstLetter += 1
            else:
                colLetter += 1

        row += 1

    #Checking for missing rows
    size1 = len(sh1)
    size2 = len(sh2)

    if size1 > size2:
        print("File2 has missing rows from ", size2+1, " to ", size1)
        missing_rows(size2, size1, sh1, cells)

    elif size2 > size1:
        print("File1 has missing rows from ", size1+1, " to ", size2)
        missing_rows(size2, size1, sh2, cells)

    print("Number of differences: ", local_counter)
    return cells


####################################################

def missing_rows(minSize, maxSize, sh, cells): #adding cells from missing rows to cells array, this function will only be called when there are missing rows in a sheet

    while minSize < maxSize:

        col = 0
        colLetter = ord("A")
        colFirstLetter = ord("@")
        sizeCol = len(sh[minSize])

        while col < sizeCol:

            if sh[minSize][col] != '':
                if colFirstLetter == ord('@'):
                    cellName = chr(colLetter) + str(minSize + 1)
                else:
                    cellName = chr(colFirstLetter) + chr(colLetter) + str(minSize + 1)

                cells.append(cellName)

            col += 1
            if colLetter == ord("Z"):
                colLetter = ord("A")
                colFirstLetter += 1
            else:
                colLetter += 1

        minSize += 1

####################################################

def highlighting_different_cells(sheet, cells):
    sheet.format(cells, {
        'backgroundColor': {
            'red': 80,
            'green': 0,
            'blue': 0
        }
    })

####################################################

print("Opening files...")
file1 = client.open("file_1")
file2 = client.open("file_2")

sheet1 = file1.get_worksheet(0)
sheet2 = file2.get_worksheet(0)
cellsFromSheet1 = compare_sheets(sheet1.get_values(), sheet2.get_values(), "sheet 1", Fore.LIGHTBLUE_EX)

sheet3 = file1.get_worksheet(1)
sheet4 = file2.get_worksheet(1)
cellsFromSheet2 = compare_sheets(sheet3.get_values(), sheet4.get_values(), "sheet 2", Fore.LIGHTGREEN_EX)

sheet5 = file1.get_worksheet(2)
sheet6 = file2.get_worksheet(2)
cellsFromSheet3 = compare_sheets(sheet5.get_values(), sheet6.get_values(), "sheet 3", Fore.LIGHTMAGENTA_EX)

sheet7 = file1.get_worksheet(3)
sheet8 = file2.get_worksheet(4)
cellsFromSheet4 = compare_sheets(sheet7.get_values(), sheet8.get_values(), "sheet 4", Fore.LIGHTYELLOW_EX)

print("Number of differences (all sheets): ", global_counter)
print(Fore.LIGHTWHITE_EX + "\nDo you want to make worksheets copies with differences highlighted? Press y or n (may take a minute)")
choice = input().lower()

if choice == 'y':
    print("Adding worksheets...")
    print("Cells with different values will be marked in dark red background in file_1")
    highlighting_different_cells(sheet1.duplicate(), cellsFromSheet1)
    #highlighting_different_cells(sheet2.duplicate(), cellsFromSheet1)
    highlighting_different_cells(sheet3.duplicate(), cellsFromSheet2)
    #highlighting_different_cells(sheet4.duplicate(), cellsFromSheet2)
    highlighting_different_cells(sheet5.duplicate(), cellsFromSheet3)
    #highlighting_different_cells(sheet6.duplicate(), cellsFromSheet3)
    highlighting_different_cells(sheet7.duplicate(), cellsFromSheet4)
    #highlighting_different_cells(sheet8.duplicate(), cellsFromSheet4)
    #file2 copies aren't made by default, if you want you can uncomment them
    print("Finished")
else:
    print("Finished")