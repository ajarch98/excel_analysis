from excel_core import duplicates_check, none_check,\
    output_df
import pandas as pd
import datetime
import os

# LOC = locations of file to use
# Place primary file in A_LOC
FILES_LOC = r"C:\Users\PAVILION\Documents\repos\excel_scripts\files"
A_FILE = "SOSI.xlsx"  # primary file
B_FILE = "NYUMBA.xlsx"  # secondary file
A_LOC = os.path.join(FILES_LOC, A_FILE)
B_LOC = os.path.join(FILES_LOC, B_FILE)

# Row number of Header (starts from 1)
A_HEADER_ROW = 1
B_HEADER_ROW = 1

# Header name of column to checks
A_COL_NAME = "LPO"
B_COL_NAME = "LPO"

# files stored in folder with syntax
# Hour-Minute-Second.date-month-year
# Change RESULTS_PATH to change location of stored data
RESULTS_PATH = r"C:\Users\Pavilion\Documents\repos\excel_repo\results"
today_date = datetime.datetime.now()
today_day = today_date.strftime("%H-%M-%S.%d-%B-%Y")
FOLDER_PATH = os.path.join(RESULTS_PATH, today_day)
os.mkdirs(FOLDER_PATH)


def check_dfs_and_print(
        df, comp_df,
        df_col_name, comp_df_col_name,
        matches_file, non_matches_file):
    """Output matched and unmatched values in df and comp_df.

    Check against df_col_name for df and comp_df_col_name for comp_df.
    Output matched values to matches_file and unmatched_values to
    non_matches_file."""
    values = comp_df[comp_df_col_name].values.astype(str)
    df["check"] = df[df_col_name].map(lambda x: str(x) in values)
    df.sort_values(by=df_col_name)

    matched_df = df[df["check"] == True]
    matched_df = matched_df.drop(["check"], axis=1)
    output_df(df, matches_file)

    unmatched_df = df[df["check"] == False]
    unmatched_df = unmatched_df.drop(["check"], axis=1)
    output_df(df, non_matches_file)


def compare_dfs(
        a_df, b_df,
        a_col_name, b_col_name,
        folder_path,
        a_file, b_file
        ):
    """Compare values in COL_NAME between A_LOC and B_LOC.

    Output matched values to MATCHES_FILE.
    Output non-matched values to NON_MATCHES_FILE.
    """
    def output_log():
        """Write log for script to LOG_FILE."""
        log_file = os.path.join(folder_path, 'log.txt')
        log_str = (
            "Contents:\n"
            f"Input file 1: {a_file}\n"
            f"Input file 2: {b_file}\n"
            "\n"
            f"Rows where {a_col_name} in {a_file} is NONE: {null_file_a}\n"
            f"Rows where {b_col_name} in {b_file} is NONE: {null_file_b}\n"
            "\n"
            f"Rows where {a_col_name} in {a_file} is duplicated: {duplicates_file_a}\n"
            f"Rows where {b_col_name} in {b_file} is duplicated: {duplicates_file_b}\n"
            "\n"
            f"Rows where {a_col_name} has matches in "
            f"{b_file} {b_col_name}: {matches_file_a}\n"
            f"NB: if {matches_file_a} does not exist, no matches were found\n"
            f"Rows where {b_col_name} has matches in "
            f"{a_file} {a_col_name}: {matches_file_b}\n"
            f"NB: if {matches_file_b} does not exist, no matches were found\n"
            "\n"
            f"Rows where {a_col_name} does not have matches in "
            f"{b_file} {b_col_name}: {non_matches_file_a}\n"
            f"NB: if {non_matches_file_a} does not exist, "
            "no empty values were found.\n"
            f"Rows where {b_col_name} does not have matches in "
            f"{b_file} {a_col_name}: {non_matches_file_b}\n"
            f"NB: if {non_matches_file_b} does not exist, "
            "no empty values were found."
            )
        with open(log_file, 'w+') as f:
            f.write(log_str)

    matches_file_a = os.path.join(folder_path, f'matches_{a_file}')
    matches_file_b = os.path.join(folder_path, f'matches_{b_file}')

    non_matches_file_a = os.path.join(folder_path, f'non_matches_{a_file}')
    non_matches_file_b = os.path.join(folder_path, f'non_matches_{b_file}')

    duplicates_file_a = os.path.join(folder_path, f'duplicates_{a_file}')
    duplicates_file_b = os.path.join(folder_path, f'duplicates_{b_file}')

    null_file_a = os.path.join(folder_path, f'null_{a_file}')
    null_file_b = os.path.join(folder_path, f'null_{b_file}')

    a_df = none_check(a_df, a_col_name, null_file_a)
    b_df = none_check(b_df, b_col_name, null_file_b)
    
    a_df = duplicates_check(a_df, a_col_name, duplicates_file_a)
    b_df = duplicates_check(b_df, b_col_name, duplicates_file_b)

    check_dfs_and_print(
        a_df, b_df,
        a_col_name, b_col_name,
        matches_file_a, non_matches_file_a)
    check_dfs_and_print(
        b_df, a_df,
        b_col_name, a_col_name,
        matches_file_b, non_matches_file_b)
    output_log()


if __name__ == "__main__":
    a_df = pd.read_excel(A_LOC, header=A_HEADER_ROW-1)
    b_df = pd.read_excel(B_LOC, header=B_HEADER_ROW-1)

    compare_dfs(
        a_df, b_df,
        A_COL_NAME, B_COL_NAME,
        FOLDER_PATH,
        A_FILE, B_FILE
        )
