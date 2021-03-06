"""Define functions for excel file usage."""


def duplicates_check(df, col_name, duplicates_file):
    """Detect duplicate values in col_name values.

    Ouput rows with duplicate values to duplicates_file.
    Return dataframe with duplicates removed.
    """
    duplicates_df = df[df.duplicated(subset=col_name, keep=False)]
    duplicates_df = duplicates_df.sort_values(by=col_name)
    output_df(duplicates_df, duplicates_file)

    non_duplicates_df = df.drop_duplicates(subset=col_name, keep=False)
    return non_duplicates_df


def none_check(df, col_name, null_file):
    """Detect None values in col_name.

    Output rows with flagged values to null_file.
    """
    df[col_name] = df[col_name].map(lambda x: x.strip() if isinstance(x, str) else x)
    none_df = df[(df[col_name].isna()) | (df[col_name] == "")]
    output_df(none_df, null_file)
    cleaned_df = df[(df[col_name].astype(str).str.len() > 0) & (df[col_name].notna())]
    return cleaned_df


def output_df(df, file):
    """Output df to file if df is not empty."""
    if not df.empty:
        df.to_excel(file)
