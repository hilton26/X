# # Reg 28 and Reg 30 Classifications
# ### Dependencies: issuers_1, issuers_2, issuers_3
# ### P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm "classifier" tab

import time
import os
import sys
from datetime import datetime
import shutil  # to copy issuers_2 and issuers_3
import subprocess
from utilities import timediff
from constants import issuers_1

start_time_r28 = time.time()
print("Executing issuer_1.py ...", "\n")
subprocess.run([sys.executable, issuers_1])

print(f"{timediff(start_time_r28, time.time())} executing issuers_1.py")
