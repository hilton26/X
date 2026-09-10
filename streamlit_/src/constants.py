#!/usr/bin/env python
# coding: utf-8

# constants
# constants from other notebook - https://stackoverflow.com/questions/6343330/importing-a-long-list-of-constants-to-a-python-file

# local folders
import os
from pathlib import Path

pth_dl = os.path.join(Path.home(), "Downloads")  # or Path.home() / "Downloads"
pthLOCAL = os.path.join(Path.home(), "Documents", "DervFiles")

# network folders
pthPIM = r"\\PIM-CPT-FS.prescient.local\PIM-Documents$"
pthW = pthPIM + r"\Working Folders\Hilton\W"  # path to working folder
pthBESA = pthPIM + r"\Shared Folders\Data\MTM"  # set folder path for bond market data
pthTest = pthPIM + r"\Working Folders\Hilton\W\Reg_Tests"
pthCmp = pthPIM + r"\Investment Operations\GRC\Compliance"
pthHdg = pthCmp + r"\PGF UCITS Share Class Hedges"
pthEXPORTS = pthCmp + r"\Derivative Cover"
pthReports = pthCmp + r"\Reg28 and Reg30 Reporting"
pthMandates = pthCmp + r"\Client Mandates"  # path to mandate summaries
pthOverdrafts = pthCmp + r"\Overdrafts"
pthDaily = pthCmp + r"\Daily"
pthMed = pthCmp + r"\Medical Schemes"
pthPtl = pthCmp + r"\Breaches\Portal"
pth_BX = pthCmp + r"\Mcaps"
pth_m_reports = pthCmp + r"\Reporting Requirements\Monthly Reports"
pthClients = pthPIM + r"\Investment Operations\Segregated Clients\Active Clients"
pth_EC = (
    pthClients
    + r"\Export Credit Insurance Corporation\ECICBAL\Reporting\Monthly Reports"
)
pth_PPSBAL = pthCmp + r"\Reporting Requirements\PIM PPSBAL"
fldr_PManco = r"V:\Operations\Manco\Prescient Manco\Reg 30 Reports"
msci_zips = pthCmp + r"\Reporting Requirements\PIM PPSBAL\Zips"

# files
pthPy = pthCmp + r"\Daily\py_reports.xlsm"
pth_r28_lmts = pthW + r"\!Reg28Templates.xlsx"
pth_tbl2_tmpl = pthW + r"\!Reg28_Tbl2.xlsx"
pth_schib_tmpl = pthW + r"\!Reg28 SchIB.xlsm"
pthCLNs = pthCmp + r"\MCaps\CLNs.xlsx"
pth_me17 = pthCmp + r"\Reporting Requirements\Monthly Reports\MonthEnd17.xlsm"
frcv_file = pthCmp + r"\Daily\Free Cover.xlsm"
pth_struct = pthPIM + r"\Working Folders\Hilton\W\structures.xlsm"
pthSttlmnt = pthCmp + r"\Daily\fund_codes.xlsx"
pthMedCirc06 = pthMed + r"\20220117 MSA Circ6 Categorisations.xlsx"
pthMedCirc12 = pthMed + r"\20230317 MSA Circ12 Categorisations.xlsx"
pthMedCirc11 = pthMed + r"\20250225 MSA Circ11of2024 Categorisations.xlsx"
pthMedCirc = pthMed + r"\20250224 MSA Circ3of2025 Categorisations.xlsx"
pthDvSmry = pthCmp + r"\Daily\derv_summary.xlsx"
pth_hdg_tmpl = pthHdg + r"\PGF Share Class Hedges.xlsm"
pth_instr = pthCmp + r"\Portal\Instruments.xlsx"
yll = pthTest + r"\yall.xlsx"
mergd = pthTest + r"\merged.xlsx"
iss_1 = pthTest + r"\issuers_1.xlsx"
iss_2 = pthTest + r"\issuers_2.xlsx"
iss_3 = pthTest + r"\issuers_3.xlsx"

# htmls
pthPrm = r"https://prime.prescient.co.za"
ptl_login = pthPrm + r"/portfolio-management/compliance"
ptl_b_rpt = ptl_login + r"/breach-report"  # breach report page
ptl_b_rpt_n = ptl_login + r"/breach-report-new"  # breach report new page
eagle_default = r"https://eagleportal.prescient.co.za/Default.aspx"
eagle_root = r"https://eagleportal.prescient.co.za/Queries/Query.aspx?rpt="
prime_data = r"https://prime.prescient.co.za/data-services"
jse_data = prime_data + r"/data-browsers/jse-index-constituents"
credit_meta = prime_data + r"/data-loggers/credit-meta"

# scripts folder
pth_gitrepo = r"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo"
selenium_drivers = r"C:/SeleniumDrivers"
pre_issuers_1 = pth_gitrepo + r"/gemsmed_pre_issuers_1.py"
issuers_1 = pth_gitrepo + r"/issuers_1.py"
issuers_2 = pth_gitrepo + r"/issuers_2.py"
issuers_3 = pth_gitrepo + r"/issuers_3.py"
issuers_1_nb = pth_gitrepo + r"/issuers_1.ipynb"
issuers_2_nb = pth_gitrepo + r"/issuers_2.ipynb"
issuers_3_nb = pth_gitrepo + r"/issuers_3.ipynb"
const = pth_gitrepo + r"/constants.py"
dc_do = pth_gitrepo + r"/derv_checker_downloading.py"
dc_co = pth_gitrepo + r"/derv_checker_compiling.py"
dc_su = pth_gitrepo + r"/derv_checker_summarising.py"
dc_fr = pth_gitrepo + r"/derv_checker_freecover.py"
pg_do = pth_gitrepo + r"/pgf_downloading.py"
pg_co = pth_gitrepo + r"/pgf_compiling.py"
dc_do_nb = pth_gitrepo + r"/derv_checker_downloading.ipynb"
dc_co_nb = pth_gitrepo + r"/derv_checker_compiling.ipynb"
dc_su_nb = pth_gitrepo + r"/derv_checker_summarising.ipynb"
dc_fr_nb = pth_gitrepo + r"/derv_checker_freecover.ipynb"
pg_do_nb = pth_gitrepo + r"/pgf_downloading.ipynb"
pg_co_nb = pth_gitrepo + r"/pgf_compiling.ipynb"
prp = pth_gitrepo + r"/prp.py"

# cave
import pandas as pd

df = pd.read_excel(pthPy, sheet_name="creds", usecols="J", header=None).dropna()
p_al = df.iloc[0, 0]
p_xe = df.iloc[1, 0]

# numerical constants
dervcoverthreshold = 10  # cut-off value  below which funds must be presented

# constants - supercats and superhens for use in function fghi()
# iterating through a nested dictionary - https://www.programiz.com/python-programming/nested-dictionary
supercats = {
    "3(f)": {
        "3(f) 2.1(e)(ii)": ["2.1(e)(ii)"],
        "3(f) 3.1(b)": ["3.1(b)"],
        "3(f) 4.1(b)": ["4.1(b)"],
        "3(f) 8": ["8.1(a)(i)", "8.1(a)(ii)", "8.2(a)(i)", "8.2(a)(ii)"],
        "3(f) 9": ["9.1(a)(i)", "9.1(a)(ii)", "9.2(a)(i)", "9.2(a)(ii)"],
        "3(f) 10": ["10.1", "10.2"],
    },
    "3(g)": {
        "3(g) 3.1(b)": ["3.1(b)"],
        "3(g) 9": ["9.1(a)(i)", "9.1(a)(ii)", "9.2(a)(i)", "9.2(a)(ii)"],
    },
    "3(h)": {
        "3(h) 1.1": ["1.1(a)", "1.1(b)", "1.1(c)", "1.1(d)"],
        "3(h) 2.1(c)": ["2.1(c)(i)", "2.1(c)(ii)", "2.1(c)(iii)", "2.1(c)(iv)"],
    },
    "3(i)": {
        "3(i) cash": ["1.2(a)", "1.2(b)", "1.2(c)"],
        "3(i) debt": [
            "2.1(b)",
            "2.2(a)",
            "2.2(a)(i)",
            "2.2(a)(ii)",
            "2.2(a)(iii)",
            "2.2(a)(iv)",
            "2.2(b)(i)",
            "2.2(b)(ii)",
            "2.2(c)(i)",
            "2.2(c)(ii)",
            "2.2(d)(i)",
            "2.2(d)(ii)",
            "2.2(e)(i)",
            "2.2(e)(ii)",
        ],
        "3(i) equity": ["3.2(a)(i)", "3.2(a)(ii)", "3.2(a)(iii)", "3.2(b)"],
        "3(i) property": ["4.2(a)(i)", "4.2(a)(ii)", "4.2(a)(iii)", "4.2(b)"],
        "3(i) commodities": ["5.2(a)(i)", "5.2(a)(ii)"],
        "3(i) hpefs": ["8.2(a)(i)", "8.2(a)(ii)", "9.2(a)(i)", "9.2(a)(ii)", "10.2"],
    },
}

# superhens for single digit and triple digit sub-totals
superhens = {
    "1": {
        "1.1": {"1.1(a)": {}, "1.1(b)": {}, "1.1(c)": {}, "1.1(d)": {}},
        "1.2": {"1.2(a)": {}, "1.2(b)": {}, "1.2(c)": {}},
    },
    "2": {
        "2.1": {
            "2.1(a)": {},
            "2.1(b)": {},
            "2.1(c)": {"2.1(c)(i)", "2.1(c)(ii)", "2.1(c)(iii)", "2.1(c)(iv)"},
            "2.1(d)": {"2.1(d)(i)", "2.1(d)(ii)"},
            "2.1(e)": {"2.1(e)(i)", "2.1(e)(ii)"},
        },
        "2.2": {
            "2.2(a)": {"2.2(a)(i)", "2.2(a)(ii)", "2.2(a)(iii)", "2.2(a)(iv)"},
            "2.2(b)": {},
            "2.2(c)": {"2.2(c)(i)", "2.2(c)(ii)", "2.2(c)(iii)", "2.2(c)(iv)"},
            "2.2(d)": {"2.2(d)(i)", "2.2(d)(ii)"},
            "2.2(e)": {"2.2(e)(i)", "2.2(e)(ii)"},
        },
    },
    "3": {
        "3.1": {"3.1(a)": {"3.1(a)(i)", "3.1(a)(ii)", "3.1(a)(iii)"}, "3.1(b)": {}},
        "3.2": {"3.2(a)": {"3.2(a)(i)", "3.2(a)(ii)", "3.2(a)(iii)"}, "3.2(b)": {}},
    },
    "4": {
        "4.1": {"4.1(a)": {"4.1(a)(i)", "4.1(a)(ii)", "4.1(a)(iii)"}, "4.1(b)": {}},
        "4.2": {"4.2(a)": {"4.2(a)(i)", "4.2(a)(ii)", "4.2(a)(iii)"}, "4.2(b)": {}},
    },
    "5": {
        "5.1": {"5.1(a)": {"5.1(a)(i)", "5.1(a)(ii)"}},
        "5.2": {"5.2(a)": {"5.2(a)(i)", "5.2(a)(ii)"}},
    },
    "6": {"6(a)": {}, "6(b)": {}},
    "7": {},
    "8": {
        "8.1": {"8.1(a)": {"8.1(a)(i)", "8.1(a)(ii)"}},
        "8.2": {"8.2(a)": {"8.2(a)(i)", "8.2(a)(ii)"}},
    },
    "9": {
        "9.1": {"9.1(a)": {"9.1(a)(i)", "9.1(a)(ii)"}},
        "9.2": {"9.2(a)": {"9.2(a)(i)", "9.2(a)(ii)"}},
    },
    "10": {"10.1": {}, "10.2": {}},
}


# (1) eagle report types, their short codes, and their URLs
# eagle_root = r"https://eagleportal.prescient.co.za/Queries/Query.aspx?rpt="
report_types_dict = {
    "r28i": [
        "Reg 28 Report - Incl Effective Exposure",
        eagle_root + "Reg28withExposure",
    ],
    "parn": ["Portfolio Analytics Report - New", eagle_root + "PortfolioAnalytics"],
    "derv": ["Derivative Exposure", eagle_root + "DerivativeExposure"],
    "trad": ["Trades Report", eagle_root + "TRANSACTION"],
    "scty": ["Security Cross Reference", eagle_root + "SecurityCrossRef"],
    "dflw": ["Daily Flows", eagle_root + "FLOWS"],
    "utps": ["Unit Trust Prices", eagle_root + "UTPRICES"],
    "fnav": ["Fund Net Asset Value", eagle_root + "NetAsset"],
    "tcrf": ["Trades Cross Reference", eagle_root + "TRADES%20REFERENCE"],
    "cact": ["Cash Activity Details", eagle_root + "CSHACTIVITY"],
}

# dictionary for unzipping MSCI index data
msci_dict = {
    "msci_wo": "WORLD SELECTION",
    "msci_em": "EM SELECTION",
    "msci_ac": "AC WORLD",
    "msci_gl": "THE WORLD INDEX",
}
