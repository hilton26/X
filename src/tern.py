#!/usr/bin/env python
# coding: utf-8

# Ad-hoc Portfolio Analytics Report (PARN) download from Eagle Portal.
# Edit `funds` and `rpt_date` before running; the file lands in Downloads.

import time
import os
from datetime import datetime
from constants import pth_dl
from utilities import timediff, osprey

start_time = time.time()

funds    = "PIMBAL,GTCWP2,HOLEQU"
rpt_date = datetime(2026, 4, 9)
name     = "PIMBAL_GTCWP2_HOLEQU"

print(f"Downloading PARN for {funds} as at {rpt_date.strftime('%d %b %Y')} ...")
osprey("parn", funds, rpt_date, rpt_date, name, "csv")
print(f"\n{timediff(start_time, time.time())} — download complete")

os.startfile(pth_dl)
