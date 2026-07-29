import requests
import urllib3
from datetime import datetime, timedelta
import csv
import pandas as pd
from constants import pthPy, pth_fx

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def get_usd_zar_rates(start_date: str, end_date: str) -> dict:
    """
    Fetch USD/ZAR exchange rates over a date range.
    Uses the free frankfurter.app API (no key required).

    Args:
        start_date: "YYYY-MM-DD"
        end_date:   "YYYY-MM-DD"

    Returns:
        dict of {date: rate} where rate is ZAR per 1 USD
    """
    url = f"https://api.frankfurter.app/{start_date}..{end_date}"
    params = {"from": "USD", "to": "ZAR"}

    response = requests.get(url, params=params, verify=False)
    response.raise_for_status()
    data = response.json()

    # Extract just the ZAR rates
    rates = {date: values["ZAR"] for date, values in data["rates"].items()}
    return dict(sorted(rates.items()))


if __name__ == "__main__":
    df = pd.read_excel(pthPy, sheet_name="arc", usecols="BG", header=0).dropna()
    dt_to = df.iloc[0, 0].date()
    dt_from = df.iloc[1, 0].date()
    # print(dt_from, dt_to)
    start = dt_from.strftime("%Y-%m-%d")
    end = dt_to.strftime("%Y-%m-%d")

    rates = get_usd_zar_rates(start, end)

    # print(f"USD/ZAR rates from {start} to {end}")
    # print(f"{'Date':<15} {'ZAR per USD':>12}")
    # print("-" * 28)
    # for date, rate in rates.items():
    #     print(f"{date:<15} {rate:>12.4f}")

    # Summary stats
    values = list(rates.values())
    # print("-" * 28)
    # print(f"{'Min':<15} {min(values):>12.4f}")
    # print(f"{'Max':<15} {max(values):>12.4f}")
    # print(f"{'Average':<15} {sum(values) / len(values):>12.4f}")

    with open(pth_fx, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["date", "usd_zar"])
        writer.writerows(rates.items())
    print(f"\nWrote {len(rates)} rows to fx_rates.csv")

    # from os import startfile
    # startfile(pth_fx)
