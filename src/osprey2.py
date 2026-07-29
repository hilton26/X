# An Eagle report lookup function, given seven parameters
def osprey(rpt_type, funds, d_from, d_to, name, sfx="csv"):
    """
    Function:
        To download a report from the online fund accounting system for a specified report type, funds, format, and dates

    Args:
        rpt_type: A report name under Queries of the Eagle online application, including r28i, parn, derv, trad, scty, dflw, utps, fnav, tcrf, and cact
        funds:    A comma-, but without spaces, separated string of fund codes, including the ones appended with "_C", .e.g., 'PABS,PPSBAL_C,SMMAIF'
        d_from:   A start date for the report in datetime format, e.g., datetime(2025,5,1)
        d_to:     An  end date for the report in datetime format, e.g., datetime(2025,5,30)
        name:     A descriptive name to be added to the downloaded report to make it more identifiable
        sfx:      A file name extension specifying the report format, i.e., 'xls' or 'csv'
        al:       The user name for the online application
        xe:       The user credential for the online application

    Returns:
        A downloaded report in the local Downloads folder renamed to identify it
    """

    # present expected file name and whether it already exists in the Downloads folder, before downloading the report
    to_date = f" to {d_to.strftime('%d%b%Y')}" if d_to != d_from else ""
    new_file_name = (
        f"{rpt_type.upper()} {name}({len(funds.split(','))}) "
        f"{d_from.strftime('%d%b%Y')}{to_date}"
    )
    print("(2) loading libraries")
    # (2) load libraries
    import time

    start_time_osprey = time.time()

    from datetime import datetime  # , timedelta
    from utilities import timediff, latest_file
    import os, re
    import pandas as pd
    from constants import pth_dl, eagle_default, report_types_dict

    file_exists = (pth_dl / f"{new_file_name}.{sfx}").exists()
    print(
        f"Expected file name based on input values:\n"
        f"   {new_file_name}.{sfx}\n"
        f"which {'already exists' if file_exists else 'does not yet exist'} in the Downloads folder.\n"
    )

    # access environment variables
    from dotenv import load_dotenv  # to access environment variables from .env file

    load_dotenv()  # take environment variables from .env file

    # selenium suite of tools
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    # from selenium.webdriver.support.select import Select
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import NoSuchElementException
    from selenium.common.exceptions import (
        NoAlertPresentException,
    )  # to handle the eagleportal.prescient.co.za alerts
    from selenium.common.exceptions import TimeoutException

    # print(funds)

    print(
        "(2a) Checking if rpt_type is 'fnav', in which case, replacing '_C' in fund name"
    )
    # (2a) for fnav rpt_type, first remove "_C" from the list of funds else the FNAV report will return "No data returned for the input criteria."
    if rpt_type == "fnav":
        sfx = "csv"
        lkup = pd.DataFrame(
            pd.Series(funds.split(",")), columns=["funds_ante"]
        )  # create a dataframe to look up before and after fund codes
        lkup["funds_post"] = lkup["funds_ante"].apply(
            lambda x: x.replace("_C", "") if x.endswith("_C") else x
        )
        funds = ",".join(lkup["funds_post"].astype(str))

    print("(3) setting report suffix for .xls vs .csv")
    # (3) set the report suffix
    t = "0" if sfx == "csv" else "4"  # report format: DXI4[0] for .xls[.csv]

    max_retries = 5
    batch_succeeded = False
    for attempt in range(1, max_retries + 1):
        print(f"(4) Importing the webdriver (attempt {attempt}/{max_retries})")
        # (4) assign the browser driver
        from selenium import webdriver

        driver = webdriver.Firefox()

        print("(5) Starting the web driver and opening the browser ...")
        # (5) open the browser on the default web page
        # eagle_default = r"https://eagleportal.prescient.co.za/Default.aspx"
        driver.get(eagle_default)  # default page
        wait = WebDriverWait(
            driver, 10
        )  # https://selenium-python.readthedocs.io/waits.html, max wait for elements to appear

        print("(6) Entering credentials to open the browser on the default web page...")
        # (6) log in
        driver.find_element(
            By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_UserName"
        ).send_keys(os.getenv("EAGLE_UN"))
        driver.find_element(
            By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_Password"
        ).send_keys(os.getenv("EAGLE_PW"))
        driver.find_element(
            By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_LoginButton"
        ).click()

        print("(7) having logged in, opening the selected report page ...")
        # (7) having logged in, open the selected report page
        report_link = report_types_dict[rpt_type][1]
        driver.get(
            report_link
        )  # a hyperlink for the report page selected in the function osprey()

        try:
            print(
                "(7a) testing for the presence of an alert and accepting it if it exists"
            )
            # (7(a)) test for the presence of an alert
            # this solution from Gemini prompt 17 Sep 2025: "python selenium test
            # for the presence of alert text"
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert.accept()  # Or alert.dismiss()
            except TimeoutException:
                # print("No alert present per TimeoutException.")
                pass
            except NoAlertPresentException:
                # print("No alert present per NoAlertPresentException.")
                pass

            print("(8) switching to the query page ...")
            # (8) switch to the query page
            driver.find_element(By.CSS_SELECTOR, "#ModifyLinkLabel").click()

            print("(9) updating the FROM calendar ...")
            # (9) update the FROM calendar
            date_selector_fr = driver.find_element(
                By.CSS_SELECTOR,
                'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_I"]',
            )  # FROM date element
            driver.execute_script(
                f'arguments[0].value = "{d_from.strftime("%#m/%#d/%Y")}";',
                date_selector_fr,
            )  # FROM date without leading zeroes
            driver.find_element(
                By.CSS_SELECTOR,
                'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_I"]',
            ).click()  # click inside FROM calendar
            driver.find_element(
                By.CSS_SELECTOR,
                'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_B-1"]',
            ).click()  # update the FROM calendar

            print("(10) updating the TO calendar, if it exists ...")
            # (10) if it exists, update the TO calendar
            try:  # https://stackoverflow.com/questions/38022658/selenium-python-handling-no-such-element-exception
                date_selector_to = driver.find_element(
                    By.CSS_SELECTOR,
                    'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_I"]',
                )  # calendar
                driver.execute_script(
                    f'arguments[0].value = "{d_to.strftime("%m/%d/%Y")}";',
                    date_selector_to,
                )
                driver.find_element(
                    By.CSS_SELECTOR,
                    'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_I"]',
                ).click()  # click inside TO calendar
                driver.find_element(
                    By.CSS_SELECTOR,
                    'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_B-1"]',
                ).click()  # update the TO calendar
            except (
                NoSuchElementException
            ):  # in the event that the selected report does not have a "to" calendar
                pass

            print(
                "(11) getting the web element for the FUND LIST and assigning values to it ..."
            )
            # (11) get the web element for the FUND LIST and assign values to it
            ct100_SelectedIds = (
                'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedIds"]'
            )
            fund_selector = driver.find_element(
                By.CSS_SELECTOR,
                ct100_SelectedIds,
            )
            driver.execute_script(f'arguments[0].value = "{funds}";', fund_selector)

            print("(12) clicking 'Entity ID' to trigger the report generation ...")
            # (12) click the table header where "Entity ID" resides
            ct100_FUND0 = 'table[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedItemsGrid_DXHeaderTable"]'
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ct100_FUND0))
            ).click()  # fund code banner
            time.sleep(5)  # arbitrary 5 second wait

            print(
                "(12a) testing for the presence of an authentication \
    alert after clicking Submit ..."
            )
            # (12a) test for the presence of an authentication alert after submitting the report query
            try:
                WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert.send_keys(
                    os.getenv("EAGLE_UN") + Keys.TAB + os.getenv("EAGLE_PW")
                )
                alert.accept()
            except TimeoutException:
                pass
            except NoAlertPresentException:
                pass

            print(
                "(13) getting the web element of the 'Submit' button and then clicking it ..."
            )
            # (13) get the web element of the 'Submit' button and then click it
            submit_button = driver.find_element(
                By.CSS_SELECTOR,
                'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_RunBtn"]',
            )
            submit_button.click()

            # added by Claude Code Interpreter on 17 Sep 2025 in response to
            # "python selenium test for the presence of alert text"
            print(
                "(13a) testing for the presence of an authentication \
    alert after clicking Submit ..."
            )
            # (13a) test for the presence of an authentication alert after submitting the report query
            try:
                WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert.send_keys(
                    os.getenv("EAGLE_UN") + Keys.TAB + os.getenv("EAGLE_PW")
                )
                alert.accept()
            except TimeoutException:
                pass
            except NoAlertPresentException:
                pass

            print("(14) waiting for and then clicking the export button ...")
            # (14) Wait for and then click the export button and then the xls download button
            # https://stackoverflow.com/questions/56085152/selenium-python-error-element-could-not-be-scrolled-into-view
            start_time_14 = time.time()
            WebDriverWait(driver, 1000).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[id="DistrBtn"]'))
            ).click()
            print(
                f"  Waited {timediff(start_time_14, time.time())} for \
    the export format button to be clickable and then clicked it ..."
            )

            start_time_14a = time.time()
            print("(14a) Waiting for and then clicking the export format button ...")
            # (14a) Waiting for and then clicking the export format button ...
            WebDriverWait(driver, 1000).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, f'td[id="ExportMnu_DXI{t}_T"]')
                )
            ).click()
            print(
                f"  Waited {timediff(start_time_14a, time.time())} for \
    the download button to be clickable and then clicked it ..."
            )

            print("(14b) waiting for the download to complete ...")
            # (14b) waiting for the download to complete ...
            start_time_dl = time.time()
            t_dl = 30  # downloaded within the past t_dl seconds
            tz = 60  # max wait time of 1 minutes for download to complete and
            # moving on to the next batch, to avoid getting stuck
            # on a batch if the download gets stuck for some reason
            dl_pattern = re.compile(
                rf"{re.escape(report_types_dict[rpt_type][0])}.*\.{sfx}$"
            )
            while True:
                candidates = [
                    f
                    for f in pth_dl.iterdir()
                    if dl_pattern.search(f.name)
                    and time.time() - f.stat().st_mtime <= t_dl
                ]  # report type, suffix and downloaded within the past 30 seconds
                if not list(pth_dl.glob("*.part")) and candidates:
                    break
                if time.time() - start_time_dl > tz:
                    print(
                        f"  Warning: timed out waiting for download after {tz} seconds"
                    )
                    break
                time.sleep(1)
            print(f"  Download completed after {timediff(start_time_dl, time.time())}")

            print("(14c) renaming the downloaded file ...")
            #   (14c) rename the downloaded file
            folder_path = str(pth_dl)

            latest_file(folder_path, sfx, new_file_name)

            print(f"  Renamed to: {new_file_name}")

            if rpt_type == "fnav":
                filen = os.path.join(folder_path, new_file_name + f".{sfx}")
                fnav = pd.read_csv(filen)
                fnav = fnav.merge(
                    lkup, how="left", left_on="NAV Entity ID", right_on="funds_post"
                )
                fnav["NAV Entity ID"] = fnav["funds_ante"]
                fnav.drop(columns=["funds_ante", "funds_post"], inplace=True)
                fnav.to_csv(filen, index=False)

            print("(14d) closing the web driver ...")
            # (14d) close the driver now that the file is fully downloaded and renamed
            driver.quit()
            batch_succeeded = True
            break  # success — exit the retry loop

        except Exception as e:
            print(f"Exception in steps 7a onwards: {e}")
            try:
                driver.close()
                driver.quit()
            except Exception:
                pass
            if attempt < max_retries:
                print(
                    f"  Retrying from step 4 (attempt {attempt + 1}/{max_retries}) ..."
                )
            else:
                print(f"  All {max_retries} attempts failed for attempt {max_retries}.")

    print(f"\nDownloaded {new_file_name}.{sfx}")
