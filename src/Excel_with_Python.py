# Source: https://www.youtube.com/watch?v=q6Mc_sAPZ2Y

import openpyxl
import os
from pathlib import Path

# filename = os.path.join('C','Users','hilton.netta','Documents','20210630 NFMW Funds Reg28.xlsx')
pth = Path("C/Users/hilton.netta/Documents/")
filename = pth / '20210630 NFMW Funds Reg28.xlsx'

wb = openpyxl.load_workbook('filename')
type(wb)