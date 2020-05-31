import pandas as pd
import datetime
import os
from excel_core import output_df

# LOC = location of file to use
LOC = r"C:\Users\PAVILION\Desktop\SampleData3.xlsx"
# Column number of Header (starts from 1)
HEADER_ROW = 1

# Header name of column to check
COL_NAME = "serial number"
# Length of values to check
LENGTH = 7

# files stored in folder with syntax
# Hour-Minute-Second.date-month-year
# Change DESKTOP_PATH to change location of stored data
DESKTOP_PATH = r"C:\Users\Pavilion\Desktop"
today_date = datetime.datetime.now()
today_day = today_date.strftime("%H-%M-%S.%d-%B-%Y")
folder_path = os.path.join(DESKTOP_PATH, today_day)
os.mkdirs(folder_path)
FLAGGED_FILE = os.path.join(folder_path, 'non_std_len.xlsx')
CLEANED_FILE = os.path.join(folder_path, 'std_len.xlsx')
LOG_FILE = os.path.join(folder_path, 'log.txt')


def output_log():
    """Write log to LOG_FILE for length checking."""
    log_str = ("Contents:\n"
               f"Input file: {LOC}\n"
               f"Length checked: {LENGTH}\n"
               f"Rows where {COL_NAME} is non_standard: {FLAGGED_FILE}\n"
               f"NB, if {FLAGGED_FILE} does not exist, no values with"
               f"length not of {LENGTH} were found."
               )
    with open(LOG_FILE, 'w+') as f:
        f.write(log_str)


def check_lengths(df, length):
    """
    Check lengths of values in column COL_NAME.

    Return df containing rows where column value size is not equal to LENGTH.
    """
    df["check"] = df[COL_NAME][df[COL_NAME].astype(str).str.len() != length]
    checked_df = df[df["check"].notna()]
    checked_df = checked_df.drop("check", axis=1)
    cleaned_df = df[df["check"].isna()]
    cleaned_df = cleaned_df.drop("check", axis=1)
    output_df(cleaned_df, CLEANED_FILE)
    output_df(checked_df, FLAGGED_FILE)


if __name__ == "__main__":
    df = pd.read_excel(LOC, header=HEADER_ROW-1)
    check_lengths(df, LENGTH)
    output_log()
