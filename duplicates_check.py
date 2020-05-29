import pandas as pd
import datetime
import os
from .core_excel import duplicates_check, none_check

# LOC = location of file to use
LOC = r"C:\Users\PAVILION\Desktop\NYUMBA.xlsx"
# Column number of Header (starts from 1)
HEADER_ROW = 1

# Header name of column to check
COL_NAME = "LPO"

# files stored in folder with syntax
# Hour-Minute-Second.date-month-year
# Change DESKTOP_PATH to change location of stored data
DESKTOP_PATH = r"C:\Users\Pavilion\Desktop"
today_date = datetime.datetime.now()
today_day = today_date.strftime("%H-%M-%S.%d-%B-%Y")
folder_path = os.path.join(DESKTOP_PATH, today_day)
os.mkdir(folder_path)
DUPLICATES_FILE = os.path.join(folder_path, 'duplicates.xlsx')
NON_DUPLICATES_FILE = os.path.join(folder_path, 'unique.xlsx')
NULL_FILE = os.path.join(folder_path, 'null.xlsx')
LOG_FILE = os.path.join(folder_path, 'log.txt')


def output_log():
    log_str = (
        "Contents:\n"
        f"Input file: {LOC}\n"
        f"Duplicates found: {DUPLICATES_FILE}\n"
        f"Unique rows: {NON_DUPLICATES_FILE}\n"
        f"Rows where {COL_NAME} is NULL: {NULL_FILE}\n"
        f"NB, if {NULL_FILE} does not exist, no empty values were found."
        )
    with open(LOG_FILE, 'w+') as f:
        f.write(log_str)


if __name__ == "__main__":
    df = pd.read_excel(LOC, header=HEADER_ROW-1)
    cleaned_df = duplicates_check(df, COL_NAME, DUPLICATES_FILE)
    cleaned_df.to_excel(NON_DUPLICATES_FILE)
    none_check(df, COL_NAME, NULL_FILE)
    output_log()
