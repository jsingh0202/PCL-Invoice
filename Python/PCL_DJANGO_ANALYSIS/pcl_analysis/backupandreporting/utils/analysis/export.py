import pandas as pd
from tkinter import Tk, filedialog
from .sheetdata import get_data
from decimal import Decimal, ROUND_HALF_UP


def compare_to_overhaul(df):
    """Compares the 'TCV' amounts in the DataFrame to the values from the Overhaul data.2

    Args:
        df (dataframe): DataFrame containing the export data.

    Returns:
        Dataframe: DataFrame containing rows where the 'TCV' amount does not match the Overhaul data.
    """
    data: dict = get_data()

    df2 = df.copy()
    nan = df2[df2.isnull().any(axis=1)]
    df2.drop(nan.index, inplace=True)

    df2["Expected Value"] = df2["Description"].apply(
        lambda desc: data.get(desc, Decimal("NaN"))
    )

    invalid = df2[
        df2.apply(
            lambda row: Decimal(str(row["Total Contract Value"]))
            .quantize(Decimal("0.01"))
            .normalize()
            != row["Expected Value"],
            axis=1,
        )
    ]

    cols = list(invalid.columns)
    tv_index = cols.index("Total Contract Value")
    ev_index = cols.index("Expected Value")
    cols.insert(tv_index + 1, cols.pop(ev_index))

    return invalid[cols]


def html_df(df):
    """
    Converts a DataFrame to an HTML table with Bootstrap classes.

    Args:
        df (DataFrame): The DataFrame to convert.

    Returns:
        str: HTML string of the DataFrame.
    """
    return df.to_html(
        classes="table table-bordered table-hover table-striped table-sm", index=False
    )


def get_duplicates(df, col):
    """
    Helper Function for getting duplicate values within a dataframe.

    Args:
        df (dataframe): dataframe to check.
        col (String): The name of the column to be checked.

    Returns:
        df: Df of the export with rows where the column is duplicated.
    """
    df2 = df.copy()
    nan = df2[df2.isnull().any(axis=1)]
    df2.drop(nan.index, inplace=True)
    c = df2[col]

    invalid = df2[c.duplicated(keep=False)]

    return invalid


def get_small_values(df, col):
    """
    Helper Function for getting small values within a dataframe.

    Args:
        df (dataframe): dataframe to check.
        col (String): The name of the column to be checked.

    Returns:
        df: Df of the export with rows where the column is between 0 and 1.
    """
    c = df[col]
    invalid = df[(c > 0) & (c < 1) | (c < 0) & (c > -1)]

    return invalid


def get_tpd_greater_than_tcv(df):
    """
    Helper Function for getting rows where TPD is greater than TCV.

    Args:
        df (dataframe): dataframe to check.

    Returns:
        df: Df of the export with rows where TPD is greater than TCV.
    """
    tpd = df["Total Progress to Date"]
    tcv = df["Total Contract Value"]

    invalid = df[tpd > tcv]

    return invalid


def calculate_invalid(df, x, y, a, b, op):
    """
    Helper Function for calculating invalid values within a dataframe.

    Args:
        df (dataframe): dataframe to check.
        x (String): The name of calculated column created to check against y.
        y (String): The name of the column being verified.
        a (String): The name of the first column to be used in the calculation.
        b (String): The name of the second column to be used in the calculation.
        op (Lambda): The operation to be performed on the two columns a and b.

    Returns:
        df: Df of the export with rows where the calculated column does not match the expected value.
    """
    # copy df and create calculated column
    df2 = df.copy()
    df2[x] = op(df2[a], df2[b])

    #  check for invalid values
    mismatch = (df2[y] - df2[x]).abs() > 1e-2
    invalid = df2[mismatch].copy()

    #  insert calculated column next to the expected column
    cols = list(invalid.columns)
    y_index = cols.index(y)
    cols.insert(y_index + 1, cols.pop(cols.index(x)))

    return invalid[cols]


def check_percent_complete(df):
    """
    Checks for rows where the % Complete is not between 0 and 100.

    Args:
        df (dataframe): Df of the export.

    Returns:
        df: Df of the export with rows where % Complete is not between 0 and 100.
    """
    perc_complete = df["% Complete"]
    invalid = df[(perc_complete < 0) | (perc_complete > 1)]
    return invalid


def check_nan(df):
    """
    Checks for rows where NaN values are present.
    Drops the rows from the df.

    Args:
        df (dataframe): Df of the export.

    Returns:
        df: Df of the export with rows where NaN values are present.
    """
    nan = df[df.isnull().any(axis=1)]
    # df.drop(nan.index, inplace=True)
    return nan


def check_empty_description(df):
    """
    Checks for rows where the Work Release # number is not present.
    Drops the rows from the df.

    Args:
        df (dataframe): Df of the export.

    Returns:
        df: Df of the export with rows where the Work Release # number is not present.
    """
    missing = df[df["Description"].isnull()]
    # df.drop(missing.index, inplace=True)
    return missing


def select_file():
    """
    Creates a file dialog to select an Excel file.

    Raises:
        Exception: If no file is selected.

    Returns:
        String: Filepath of the export.
    """
    Tk().withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    if not file_path:
        raise Exception("No file selected.")
    return file_path


def analyze(file_path):
    """
    Analyzes the export file generated by the backup-generation.py script.
    """
    df = pd.read_excel(file_path, dtype={"Description": str})

    print("Loaded file:", file_path)

    missing = check_empty_description(df)
    nan = check_nan(df)
    invalid_perc_complete = check_percent_complete(df)
    invalid_tpd_perc_complete = calculate_invalid(
        df,
        "Calculated TPD (TCV * %C)",
        "Total Progress to Date",
        "Total Contract Value",
        "% Complete",
        lambda a, b: a * b,
    )
    invalid_tpd_prev_curr = calculate_invalid(
        df,
        "Calculated TPD (PB + CB)",
        "Total Progress to Date",
        "Previously Billed",
        "Current Billing",
        lambda a, b: a + b,
    )
    invalid_tcv = calculate_invalid(
        df,
        "Calculated TCV (TPD + Balance)",
        "Total Contract Value",
        "Total Progress to Date",
        "Balance",
        lambda a, b: a + b,
    )

    tpd_greater_tcv = get_tpd_greater_than_tcv(df)

    small_tcv = get_small_values(df, "Total Contract Value")
    small_tpd = get_small_values(df, "Total Progress to Date")
    small_pb = get_small_values(df, "Previously Billed")
    small_cb = get_small_values(df, "Current Billing")
    small_balance = get_small_values(df, "Balance")

    small_values = pd.concat(
        [
            small_tcv,
            small_tpd,
            small_pb,
            small_cb,
            small_balance,
        ]
    ).drop_duplicates()

    duplicated_wrs = get_duplicates(df, "Description")

    incorrect_billing = compare_to_overhaul(df)

    reporting_values = {
        "Missing Description - Rows where 'Description' is blank": html_df(missing),
        "NaN Values - Rows where 'NaN' values appear": html_df(nan),
        "Invalid % Complete - Rows where '% Complete' is either greater than 1 or less than 0": html_df(
            invalid_perc_complete
        ),
        "Invalid TPD (% Complete) - Rows where 'TPD' is not equal to TCV * % Complete": html_df(
            invalid_tpd_perc_complete
        ),
        "Invalid TPD (PB + CB) - Rows where 'TPD' is not equal to PB + CB": html_df(
            invalid_tpd_prev_curr
        ),
        "Invalid TPD - Rows where 'TPD' is greater than 'TCV'": html_df(
            tpd_greater_tcv
        ),
        "Invalid TCV - Rows where 'TCV' is not equal to TPD + Balance": html_df(
            invalid_tcv
        ),
        "Small Values - Rows with values that are very small": html_df(small_values),
        "Duplicate WRs - Rows where a WR shows up more than once": html_df(
            duplicated_wrs
        ),
        "Incorrect Billing - Rows where the 'TCV' amount does not match the agreed amount": html_df(
            incorrect_billing
        ),
    }

    # print(reporting_values)

    return reporting_values
