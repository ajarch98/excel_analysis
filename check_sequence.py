import pandas as pd
import datetime
import os

# LOC = location of file to use
LOC = r"C:\Users\PAVILION\Desktop\SampleData3.xlsx"
# Column number of Header (starts from 1)
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
SORTED_FILE = os.path.join(folder_path, 'unique.xlsx')
MISSING_FILE = os.path.join(folder_path, 'missing_nos.txt')
LOG_FILE = os.path.join(folder_path, 'log.txt')


def output_log():
    """Write log for script to LOG_FILE."""
    log_str = ("Contents:\n"
               f"Input file: {LOC}\n"
               f"Sorted values in: {SORTED_FILE}\n"
               f"Non-sequence values in {COL_NAME}: {MISSING_FILE}\n"
               f"NB, if {MISSING_FILE} does not exist,"
               "no missing values were found."
               )
    with open(LOG_FILE, 'w+') as f:
        f.write(log_str)


def sort_data(df):
    """Sort data in COL_NAME.

    Output sorted_data to SORTED_FILE.
    """
    df = df.sort_values(by=COL_NAME)
    df.to_excel(SORTED_FILE)


def check_missing(df):
    """Flag missing number in sequence of COL_NAME columns.

    Output missing numbers to MISSING_FILE.
    """
    df = df.sort_values(by=COL_NAME)
    df = df.reset_index(drop=True)
    missing = []
    try:
        for i in range(df[COL_NAME].size):
            if i == 0:
                continue
            if df[COL_NAME][i] - df[COL_NAME][i-1] != 1:
                missing.extend(list(range(df[COL_NAME][i-1] + 1, df[COL_NAME][i])))
    except TypeError:
        raise Exception("String found in column")
    if missing:
        f = open(MISSING_FILE, "w+")
        print("Missing numbers in the sequence are: ", file=f)
        print(*missing, sep="\n", file=f)
        f.close()


if __name__ == "__main__":
    df = pd.read_excel(LOC, header=HEADER_ROW-1)
    sort_data(df)
    check_missing(df)
