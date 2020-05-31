from excel_core import output_df, none_check
import pandas as pd
import datetime
import os

# LOC = location of file to use
LOC = r"C:\Users\PAVILION\Desktop\SampleData.xlsx"
# Row number of Header (starts from 1)
HEADER_ROW = 1

# Header name of column to check
COL_NAME = "Order Number"

# files stored in folder with syntax
# Hour-Minute-Second.date-month-year
# Change DESKTOP_PATH to change location of stored data
DESKTOP_PATH = r"C:\Users\Pavilion\Desktop"
today_date = datetime.datetime.now()
today_day = today_date.strftime("%H-%M-%S.%d-%B-%Y")
folder_path = os.path.join(DESKTOP_PATH, today_day)
os.mkdirs(folder_path)
ALNUM_FILE = os.path.join(folder_path, 'alnum.xlsx')
NULL_FILE = os.path.join(folder_path, 'null.xlsx')
LOG_FILE = os.path.join(folder_path, 'log.txt')


def output_log():
    """Ouput script info to LOG_FILE."""
    log_str = ("Contents:\n"
               f"Input file: {LOC}\n"
               f"Rows where {COL_NAME} is alphanumeric : {ALNUM_FILE}\n"
               f"NB: if {ALNUM_FILE} does not exist, no alphanumeric values were found\n"
               f"Rows where {COL_NAME} is NULL: {NULL_FILE}\n"
               f"NB: if {NULL_FILE} does not exist, no empty values were found."
               )
    with open(LOG_FILE, 'w+') as f:
        f.write(log_str)


def alnum_check(df):
    """Flag non-decimal values in COL_NAME.

    Output rows with non-decimal value in COL_NAME to ALNUM_FILE.
    """
    df["Check"] = df[COL_NAME].astype(str).str.isdecimal()
    flagged_df = df[df["Check"] == False]
    flagged_df = flagged_df.drop(["Check"], axis=1)
    if not flagged_df.empty:
        flagged_df.to_excel(ALNUM_FILE)


if __name__ == "__main__":
    df = pd.read_excel(LOC, header=HEADER_ROW-1)
    none_check(df, COL_NAME, NULL_FILE)
    alnum_check(df)
    output_log()
