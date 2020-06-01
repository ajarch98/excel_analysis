from excel_core import output_df, none_check
import pandas as pd
import datetime
import os

# LOC = location of file to use
LOC = r"C:\Users\PAVILION\Documents\repos\excel_repo\files\SampleData.xlsx"
# Row number of Header (starts from 1)
HEADER_ROW = 1

# Header name of column to check
COL_NAME = "Order Number"

# files stored in folder with syntax
# Hour-Minute-Second.date-month-year
# Change RESULTS_PATH to change location of stored data
RESULTS_PATH = r"C:\Users\Pavilion\Documents\repos\excel_repo\results"
today_date = datetime.datetime.now()
today_day = today_date.strftime("%H-%M-%S.%d-%B-%Y")
folder_path = os.path.join(RESULTS_PATH, today_day)
os.makedirs(folder_path)
ALNUM_FILE = os.path.join(folder_path, 'alnum.xlsx')
CLEANED_FILE = os.path.join(folder_path, 'digits.xlsx')
NULL_FILE = os.path.join(folder_path, 'null.xlsx')
LOG_FILE = os.path.join(folder_path, 'log.txt')


def output_log():
    """Ouput script info to LOG_FILE."""
    log_str = ("Contents:\n"
               f"Input file: {LOC}\n"
               f"Rows where {COL_NAME} is alphanumeric : {ALNUM_FILE}\n"
               f"NB: if {ALNUM_FILE} does not exist, no alphanumeric values were found\n"
               f"Rows where {COL_NAME} is a digit : {CLEANED_FILE}\n"
               f"NB: if {CLEANED_FILE} does not exist, no digits were found\n"
               f"Rows where {COL_NAME} is NULL: {NULL_FILE}\n"
               f"NB: if {NULL_FILE} does not exist, no empty values were found."
               )
    with open(LOG_FILE, 'w+') as f:
        f.write(log_str)


def alnum_check(df, flagged_file):
    """Flag non-decimal values in COL_NAME.

    Output rows with non-decimal value in COL_NAME to ALNUM_FILE.
    """
    df["Check"] = df[COL_NAME].astype(str).str.isdecimal()
    flagged_df = df[df["Check"] == False]
    flagged_df = flagged_df.drop(["Check"], axis=1)
    output_df(flagged_df, flagged_file)

    cleaned_df = df[df["Check"] == True]
    cleaned_df = cleaned_df.drop(["Check"], axis=1)
    return cleaned_df


if __name__ == "__main__":
    df = pd.read_excel(LOC, header=HEADER_ROW-1)
    none_check(df, COL_NAME, NULL_FILE)
    cleaned_df = alnum_check(df, ALNUM_FILE)
    output_df(cleaned_df, CLEANED_FILE)
    output_log()
