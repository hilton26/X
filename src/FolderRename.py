import os
from datetime import datetime # https://www.dataquest.io/blog/python-datetime-tutorial/

dte = datetime.now()
rptdate = datetime.strptime(dte,"%Y%m%d") # time format yyyy-mm-dd

# pth = "\\pim-cpt-fs\profile$\PProfile Operations\Legal\Legal and Compliance Framework\Compliance Certificates\OM Wealth Compliance Reports\"
path = os.path.join(pth, rptdate)
if os.path.exists(path):
    print(path + ' : exists')
    if os.path.isdir(path):
        print(path + ' : is a directory')