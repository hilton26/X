def bankbal(report_date):
    import pandas as pd
    import os
    from constants import pthOverdrafts

    # if they're available, adjoin bank balances to dfSummary
    # bank_file_sa = pthOverdrafts + rf"\{rptDate.strftime('%Y%m%d')}_overdrafts_sa.xlsx"
    bank_file_sa = (
        pthOverdrafts
        + rf"\{report_date.strftime('%Y%m%d')}_unconfirmed_cash_balances_sa.xlsx"
    )

    # extract overdrafts
    if os.path.exists(bank_file_sa):
        bank = pd.read_excel(bank_file_sa, usecols="A,C,E,H", sheet_name=0, header=0)
        # "Value Date", "Fund Code",
        # "Unconfirmed Balance BNK", "Currency" on 1st sheet
        # print(bank.columns,"\n", bank.shape)

        bank = bank[bank["Currency"] == "ZAR"]
        # print(bank.columns, "\n", bank.shape)
        bank["Value Date"] = pd.to_datetime(bank["Value Date"])
        bank_cols = {"Unconfirmed Balance BNK": "Bank Bal (ZAR)"}
        bank.rename(columns=bank_cols, inplace=True)

        return bank

    else:
        columns = ["Value Date", "Fund Code", "Bank Bal (ZAR)", "Currency"]
        return pd.DataFrame(columns=columns)


# TEST bankbal()
from datetime import date
from utilities import bankbal

rptDate = date(2026, 8, 28)
print((type(rptDate), "\n", rptDate))

bank_returned = bankbal(rptDate)
print(bank_returned, "\n", bank_returned.dtypes)
