{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 48,
   "id": "6758c728",
   "metadata": {},
   "outputs": [],
   "source": [
    "from datetime import datetime\n",
    "import time\n",
    "\n",
    "def timediff(start_time, end_time):\n",
    "    format_str       = '%Y-%m-%d %H:%M:%S.%f'\n",
    "    start_datetime   = datetime.strptime(start_time.strftime(format_str), format_str)\n",
    "    if isinstance(end_time, float):\n",
    "        end_datetime = datetime.fromtimestamp(end_time)\n",
    "    else:\n",
    "        end_datetime = datetime.strptime(end_time.strftime(format_str), format_str)\n",
    "    #end_datetime    = datetime.fromtimestamp(time.time()) if  else datetime.strptime(end_time.strftime(format_str), format_str)\n",
    "    #https://www.influxdata.com/blog/how-convert-timestamp-to-datetime-in-python/#:~:text=The%20process%20of%20converting%20a,will%20handle%20the%20conversion%20effortlessly.\n",
    "    \n",
    "    time_difference  = end_datetime - start_datetime\n",
    "    days             = time_difference.days\n",
    "    hours, remainder = divmod(time_difference.seconds, 3600)\n",
    "    minutes, seconds = divmod(remainder, 60)\n",
    "    return f'{days}d {hours}hr {minutes}min {seconds:.1f}sec' # to one (:.1f) decimal"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 49,
   "id": "875a2a8b",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Example\n",
    "#start = '2022-11-03 10:12:11.3'\n",
    "#end   = '2024-05-17 14:43:20.3'\n",
    "#print(timediff(start, end))# function to convert .xls to .xlsx using win32"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "9e782240-11b4-4882-8ca5-c0a435039a82",
   "metadata": {},
   "outputs": [],
   "source": [
    "# function to convert .xls to .xlsx using win32\n",
    "def xlsToXlsx(filepathInclXls):\n",
    "    import win32com.client as win32 # library to convert xls to xlsx\n",
    "    excel               = win32.gencache.EnsureDispatch('Excel.Application')\n",
    "    excel.DisplayAlerts = False # suppress the warning dialogue\n",
    "    wb = excel.Workbooks.Open(filepathInclXls)\n",
    "    wb.SaveAs(filepathInclXls + 'x', FileFormat = 51)     # FileFormat = 51 (56) for .xlsx (.xlx) extension\n",
    "    wb.Close()\n",
    "    excel.DisplayAlerts = True  # unsuppress Excel warning dialogue\n",
    "    return print(filepathInclXls + 'x')"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.11.5"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
