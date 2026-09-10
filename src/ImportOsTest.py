import os
from datetime import datetime # https://www.dataquest.io/blog/python-datetime-tutorial/

dte = '2021-12-04'
rptdate = datetime.strptime(dte,"%Y-%m-%d") # time format yyyy-mm-dd
print(rptdate)
print(type(rptdate))