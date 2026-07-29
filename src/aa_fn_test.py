# Function to report date and selected summary sheet option
def parn_de():

    import pandas as pd
    from constants import pthPy, pth_dl
    from utilities import prior_working_day
    from datetime import datetime
    import os

    print("Getting the reporting date and names derivative files ...")

    df = pd.read_excel(pthPy, sheet_name="arc", header=None, usecols="A,E").dropna(
        subset=[0]
    )
    k = df.iloc[2, 1]
    rptDate = (
        k if isinstance(k, datetime) else prior_working_day(datetime.today())
    )  # prior working day or report date override; has type datetime()
    summ_yn = df.iloc[3, 1]
    dervthreshold = df.iloc[4, 1] * 100
    funds = df[0].iloc[1:]

    # derive holdings and derivatives file paths
    fPARN = os.path.join(
        pth_dl,
        f"PARN ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
    )

    fDE = os.path.join(
        pth_dl,
        f"DERV ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
    )

    # check if the required files have been downloaded, else
    if not os.path.exists(fPARN) or not os.path.exists(fDE):
        sys.exit(
            f"Stopping: missing expected download(s):\n"
            f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
            f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}"
        )
        
    return fPARN, fDE, funds, rptDate, summ_yn, dervthreshold

if __name__ == "__main__":
    fPARN, fDE, funds, rptDate, summ_yn, dervthreshold = parn_de()
    print("\n\n", fPARN, "\n", fDE, "\n", funds, "\n", rptDate, "\n", summ_yn, "\n", dervthreshold)
    