import xlrd
import xlsxwriter

# USERS_LOC = location of usernames_file
# USERS_COL = column of usernames in file
USERS_LOC = r"C:\Users\PAVILION\Desktop\usernames.xlsx"
USERS_COL = 2
# HR_LOC = location of HR_List file
# HR_COL = Column of usernames in file
HR_LOC = r"C:\Users\PAVILION\Desktop\HR_List.xlsx"
HR_COL = 2


# Make sure you delete file each time you run the program
# or the new results will be appended to the bottom
# of the generated file
# RESULTS_PATH = intended path for results file
RESULTS_PATH = r"C:\Users\PAVILION\Documents\repos\excel_scripts\results\name_search.xlsx"

users_wb = xlrd.open_workbook(USERS_LOC)
users_sheet = users_wb.sheet_by_index(0)
users_sheet.cell_value(0, 0)
hr_wb = xlrd.open_workbook(HR_LOC)
hr_sheet = hr_wb.sheet_by_index(0)
hr_sheet.cell_value(0, 0)

results_book = xlsxwriter.Workbook(RESULTS_PATH)
results_sheet = results_book.add_worksheet()


row_count = 2
write_row = 0
for row in range(0, hr_sheet.nrows):
    hr_name = hr_sheet.cell_value(row, HR_COL-1) or None
    if not hr_name:
        continue
    names_l = []
    rows_l = []
    for name in hr_name.split(' '):
        results_sheet.write(row_count, 0, name)
        row_count += 1
        if len(name) < 3:
            continue
        for row in range(0, users_sheet.nrows):
            users_name = users_sheet.cell_value(row, USERS_COL-1)
            if name.lower() in users_name.lower():
                results_sheet.write(row_count, 1, users_name)
                row_count += 1
