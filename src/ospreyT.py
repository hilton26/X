import os
import time
from datetime import datetime
import holidays
from dotenv import load_dotenv  # to access environment variables from .env

load_dotenv()  # to access environment variables from .env


za_holidays = holidays.ZA()
ie_holidays = holidays.IE()
us_holidays = holidays.US()
eu_holidays = holidays.ECB()
ny_holidays = holidays.NYSE()


from constants import pth_dl, eagle_default, report_types_dict
from utilities import timediff, latest_file

# osprey() inputs variables are defined in the function for testing purposes
rpt_type = "parn"
funds = "PCBF, PCCEF, PCEQTF, PCGEARF, PCGEF, PCSHQ, PEQ, PEQF, \
    PETFIP, PEYF, PFFIF, PGCBF, PGCEF, PGIPFA, PGPCEM_C, PGPCGE_C, \
    PGPGARF, PGPGBF_C, PGPGIF_C, PGPRF_C, PICPROV, PIF, PIMBAL, \
    PIMEVO, PIMIDF, PIPF, PIPFP, PLMED, PLPRNA, PMMF, POIF_C, \
    PORFIP, PPEF, PPOS, PPSBAL_C, PPSFLEX, PPSNAM, PPSRBNQ, \
    PPWEQU, PRPABF, PRPDBF, PSIF, PSILB, PSPAM, PSSPFEQU, \
    PSTIF, PTIF, QIFFGIF, SAAMCAU, SAAMINC, SAAMMOD, SABCCSHF, SCBPF, \
    SDINCPM, SILAIF, SILEQF, SISLP, SMMAIF, SMMIBF, SMMILP, SMMRRF, \
    SNPFEQU, SSLSP, STBFIP, TRFINC, TRFWLTH, UCT3YQ, UCTINC, UCTRFBAL, \
    UCTRFINC, UNISABAL, USAIPF, UWRFCON, UWRFGRO, VMPTAA_C"
d_from = datetime.strptime("2026-04-28", "%Y-%m-%d")
d_to = datetime.strptime("2026-04-28", "%Y-%m-%d")
name = "half2"
sfx = "csv"

# def osprey(rpt_type, funds, d_from, d_to, name, sfx):
"""
Download a report from the Eagle fund accounting portal.

Args:
    rpt_type: Report code — r28i, parn, derv, trad, scty, dflw, utps, fnav, tcrf, cact
    funds:    Comma-separated fund codes (no spaces), e.g. 'PABS,PPSBAL_C,SMMAIF'
    d_from:   Start date as datetime, e.g. datetime(2025, 5, 1)
    d_to:     End date as datetime, e.g. datetime(2025, 5, 30)
    name:     Descriptive label appended to the downloaded file name
    sfx:      File format — 'xls' or 'csv'

Returns:
    Downloaded report saved to Downloads folder with an identifiable name.
"""

import pandas as pd
from constants import pth_dl, eagle_default, report_types_dict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    NoAlertPresentException,
    TimeoutException,
)

eagle_un = os.getenv("EAGLE_UN")
eagle_pw = os.getenv("EAGLE_PW")
if not eagle_un or not eagle_pw:
    raise EnvironmentError(
        "EAGLE_UN and EAGLE_PW must be set in the environment or .env file"
    )

start_time_osprey = time.time()

# for fnav, strip "_C" suffix — the report returns no data for suffixed codes
if rpt_type == "fnav":
    sfx = "csv"
    lkup = pd.DataFrame(pd.Series(funds.split(",")), columns=["funds_ante"])
    lkup["funds_post"] = lkup["funds_ante"].apply(
        lambda x: x.replace("_C", "") if x.endswith("_C") else x
    )
    funds = ",".join(lkup["funds_post"].astype(str))

t = "0" if sfx == "csv" else "4"  # DXI format index: 0=csv, 4=xls
date_fmt = "%#m/%#d/%Y" if os.name == "nt" else "%-m/%-d/%Y"  # strip leading zeros
t_report = 1000  # max seconds for Eagle to generate the report (DistrBtn to appear)
t_export = 30  # max seconds for the export menu item to appear after clicking DistrBtn
t_download = 120  # max seconds for the .part file to complete after export starts
t_message = 90  # max seconds to wait for DistrBtn after "All required criteria" warning

for attempt in range(2):
    driver = None
    try:
        ff_options = webdriver.FirefoxOptions()
        ff_options.set_capability("unhandledPromptBehavior", "dismiss")
        driver = webdriver.Firefox(options=ff_options)
        driver.set_page_load_timeout(60)
        driver.get(eagle_default)

        # Dismiss Basic Auth dialog if the login page itself presents one
        try:
            WebDriverWait(driver, 10).until(EC.alert_is_present())
            driver.switch_to.alert.dismiss()
        except (TimeoutException, NoAlertPresentException):
            pass

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_UserName")
            )
        )
        driver.find_element(
            By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_UserName"
        ).send_keys(eagle_un)
        driver.find_element(
            By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_Password"
        ).send_keys(eagle_pw)
        driver.find_element(
            By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_LoginButton"
        ).click()

        # Wait for login redirect — if it times out, proceed anyway;
        # a failed login will be caught later by the ReturnUrl check
        try:
            WebDriverWait(driver, 30).until(EC.url_changes(eagle_default))
        except TimeoutException:
            pass

        try:
            driver.get(report_types_dict[rpt_type][1])
        except TimeoutException:
            pass  # page hangs on resource loads but DOM content is usable

        try:
            WebDriverWait(driver, 10).until(EC.alert_is_present())
            driver.switch_to.alert.dismiss()
        except (TimeoutException, NoAlertPresentException):
            pass

        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#ModifyLinkLabel"))
        ).click()

        date_selector_fr = driver.find_element(
            By.CSS_SELECTOR,
            'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_I"]',
        )
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            ' arguments[0].dispatchEvent(new Event("input", {bubbles: true}));'
            ' arguments[0].dispatchEvent(new Event("change", {bubbles: true}));',
            date_selector_fr,
            d_from.strftime(date_fmt),
        )
        driver.find_element(
            By.CSS_SELECTOR,
            'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_I"]',
        ).click()
        driver.find_element(
            By.CSS_SELECTOR,
            'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_B-1"]',
        ).click()

        date_selector_to = None
        try:
            date_selector_to = driver.find_element(
                By.CSS_SELECTOR,
                'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_I"]',
            )
            driver.execute_script(
                "arguments[0].value = arguments[1];"
                ' arguments[0].dispatchEvent(new Event("input", {bubbles: true}));'
                ' arguments[0].dispatchEvent(new Event("change", {bubbles: true}));',
                date_selector_to,
                d_to.strftime(date_fmt),
            )
            driver.find_element(
                By.CSS_SELECTOR,
                'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_I"]',
            ).click()
            driver.find_element(
                By.CSS_SELECTOR,
                'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_B-1"]',
            ).click()
        except NoSuchElementException:
            pass

        fund_selector = driver.find_element(
            By.CSS_SELECTOR,
            'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedIds"]',
        )
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            ' arguments[0].dispatchEvent(new Event("input", {bubbles: true}));'
            ' arguments[0].dispatchEvent(new Event("change", {bubbles: true}));',
            fund_selector,
            funds,
        )

        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'table[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedItemsGrid_DXHeaderTable"]',
                )
            )
        ).click()

        confirmed_from = driver.execute_script(
            "return arguments[0].value", date_selector_fr
        )
        confirmed_to = (
            driver.execute_script("return arguments[0].value", date_selector_to)
            if date_selector_to is not None
            else d_to.strftime(date_fmt)
        )
        confirmed_funds = driver.execute_script(
            "return arguments[0].value", fund_selector
        )
        print(
            f"[osprey] query loaded: {rpt_type} | "
            f"from={confirmed_from} | to={confirmed_to} | funds={confirmed_funds}"
        )

        # Wait for Run button to be ready rather than a fixed sleep
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_RunBtn"]',
                )
            )
        ).click()

        # # Detect "All required criteria" message — if DistrBtn appears during
        # # the inspection window, click it immediately rather than aborting
        # distr_clicked = False
        # try:
        #     WebDriverWait(driver, 5).until(
        #         EC.presence_of_element_located(
        #             (By.XPATH, "//*[contains(text(), 'All required criteria')]")
        #         )
        #     )
        #     print(
        #         f"\n  {rpt_type.upper()}: 'All required criteria have not been selected' detected."
        #         f"\n  Waiting up to {t_message}s for the report to load anyway ..."
        #     )
        #     try:
        #         WebDriverWait(driver, t_message).until(
        #             EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[id="DistrBtn"]'))
        #         ).click()
        #         distr_clicked = True
        #     except TimeoutException:
        #         if attempt >= 1:
        #             raise RuntimeError(
        #                 f"{rpt_type.upper()}: portal rejected the run — 'All required criteria have not been selected'"
        #             )
        #         print(
        #             f"\n  {rpt_type.upper()}: criteria rejection — re-logging in and retrying..."
        #         )
        #         driver.quit()
        #         driver = None
        #         continue
        # except TimeoutException:
        #     pass  # message absent — criteria were accepted

        distr_clicked = False  # NEW
        if not distr_clicked:
            # Wait for DistrBtn, session-expired redirect, or login form, whichever comes first
            WebDriverWait(driver, t_report).until(
                EC.any_of(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[id="DistrBtn"]')),
                    EC.url_contains("ReturnUrl"),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#form1")),
                )
            )
            session_expired = "ReturnUrl" in driver.current_url or bool(
                driver.find_elements(By.CSS_SELECTOR, "#form1")
            )
            if session_expired:
                if attempt >= 1:
                    raise RuntimeError(
                        f"{rpt_type.upper()}: session expired on retry — aborting"
                    )
                print(
                    f"\n  {rpt_type.upper()}: session expired — re-logging in and retrying..."
                )
                driver.quit()
                driver = None
                continue
            driver.find_element(By.CSS_SELECTOR, 'a[id="DistrBtn"]').click()
        dl_path = str(pth_dl)
        files_before = set(os.listdir(dl_path))  # snapshot before triggering export
        WebDriverWait(driver, t_export).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, f'td[id="ExportMnu_DXI{t}_T"]')
            )
        ).click()

        # Phase 1: wait for a new file to appear — .part (in progress) or
        # already complete (download finished before first poll)
        deadline_start = time.time() + 30
        while time.time() < deadline_start:
            if set(os.listdir(dl_path)) - files_before:
                break
            time.sleep(0.5)
        deadline_end = time.time() + t_download
        while time.time() < deadline_end:
            if not any(f.endswith(".part") for f in os.listdir(dl_path)):
                break
            time.sleep(1)
        else:
            raise RuntimeError(
                f"{rpt_type.upper()}: download did not complete within {t_download}s — .part file still present"
            )

        to_date = " to " + d_to.strftime("%d%b%Y") if d_to != d_from else ""
        new_file_name = f"{rpt_type.upper()} {name}({len(funds.split(','))}) {d_from.strftime('%d%b%Y')}{to_date}"

        latest_file(dl_path, sfx, new_file_name)

        if rpt_type == "fnav":
            filen = os.path.join(dl_path, new_file_name + f".{sfx}")
            fnav = pd.read_csv(filen)
            fnav = fnav.merge(
                lkup, how="left", left_on="NAV Entity ID", right_on="funds_post"
            )
            fnav["NAV Entity ID"] = fnav["funds_ante"]
            fnav.drop(columns=["funds_ante", "funds_post"], inplace=True)
            fnav.to_csv(filen, index=False)

        # return  # download complete

    except Exception:
        print(
            f"\n  {rpt_type.upper()} report not completed after {timediff(start_time_osprey, time.time())}\n"
        )
        raise
    finally:
        if driver:
            driver.quit()
