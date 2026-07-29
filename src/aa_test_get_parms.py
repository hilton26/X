import time

start_time = time.time()
from utilities import timediff, rpt_parms

py_file = "derv"
sht_name = "qin"
result = rpt_parms(col_ref=py_file, sht_name=sht_name)
print(type(result))

if isinstance(result, tuple):
    funds, rptDate = result
    print("\n\n", funds, "\n\n", rptDate, "\n\n")
else:
    print(result)

print(timediff(start_time, time.time()))
