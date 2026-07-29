# get the reporting date and list of funds to process from the py_reports.xlsm file

########## TEST
# col_ref = "derv_portfolio"  # column reference in py_reports.xlsm to get the report date and funds to process
########## TEST


def rpt_parms(col_ref="dervs"):
    import pandas as pd
    from datetime import datetime
    from constants import pthPy
    from utilities import prior_working_day

    try:
        # find the position of the column with the py script name
        qin_heads = pd.read_excel(pthPy, sheet_name="qin", header=0, nrows=0)
        py_name = qin_heads.columns.get_loc(col_ref)
        # print(f"Position of column '{col_ref}': {py_name}")
    except:
        print(f"\n\n'{col_ref}' column not found on the 'qin' sheet.\n\n")
    else:
        # find the 'fx' column
        fx = qin_heads.columns.get_loc("fx")
        # print(f"Position of column 'fx': {fx}")

    # get report funds and dates
    df = pd.read_excel(
        pthPy, sheet_name="qin", usecols=[0, py_name, fx + 2, fx + 4], header=None
    )
    df = df.dropna(subset=[df.columns[0]])
    # print(df)

    # get report date
    # k= pd.read_excel(pthPy, sheet_name="qin", usecols=[fx+2], header=None, nrows = 1).iloc[0]
    k = df.iloc[0, 2]
    rptDate = (
        k.date()
        if isinstance(k, datetime)
        else prior_working_day(datetime.today()).date()
    )  # prior working day or report date override; has type datetime()
    # print(rptDate)

    # if generic search, get the 'to' date from sheet 'qin'
    if col_ref == "eagle_gen":
        k2 = df.iloc[0, 3]
        date_to = (
            k2.date()
            if isinstance(k2, datetime)
            else prior_working_day(datetime.today()).date()
        )  # prior working day or report date override; has type datetime()
        # print(date_to)

    # delete two date columns and the 'Count --->' row
    df = df[[0, py_name]]
    df = df.drop(df.index[[0, 1]]).reset_index(drop=True)

    # print(df)

    # funds
    funds = df[df.iloc[:, -1] == 1].iloc[:, 0].reset_index(drop=True).tolist()

    return funds, rptDate


########## TEST the ouput of this function
# if __name__ == "__main__":
#     g, h = rpt_parms("derv_portfolio")
#     print(
#         "\n",
#         g,
#         "\n",
#         type(g),
#         "\n",
#         h,
#         "\n",
#         type(h),
#         "\n",
#     )
########## TEST the ouput of this function
