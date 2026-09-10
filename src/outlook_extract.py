"""'AMT1 - Extracting Required Information from you Outlook Inbox' https://www.youtube.com/watch?v=50o6RTvYIpY&t=332s

Reads Outlook inbox (or other email folder) and extracts specific information from emails"""

import win32com.client
import pandas as pd  # to capture data into a dataframe
import re  # for regular expressions
from datetime import datetime, timedelta
from utilities import prior_working_day
import os
from pathlib import Path
from tqdm import notebook, tqdm

# define the Outlook object
# MAPI (Messaging API) is an API for Windows which allows programs to become email-aware
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# print Outlook folders
# for foldewr in outlook.Folders:
#     print(folder.Name)

# define the inbox folder to be searched
# inbox = outlook.Folders("Inbox")
inbox = outlook.GetDefaultFolder(6)  # 6 refers to the inbox folder

# fetch messages from the inbox defined above
time_filter = f"[ReceivedTime] >= '{prior_working_day(datetime.today() - timedelta(days=10)).strftime("%m/%d/%Y")}'"
print(time_filter)
#messages = inbox.Items.Restrict(time_filter)
messages = inbox.Items  # without a day filter
# print(messages)

# create running lists to capture the email data iteratively
all_names = []
all_dates = []
all_company = []
all_email = []
all_body_text = []
for i, message in tqdm(enumerate(messages)):
    if message.Subject == "Report Center - Attachment Notification".lower():
        # date_time = message.LastModificationTime # extract email date and time
        # sender_name = message.SenderName
        # sender_email = message.SenderEmailAddress
        # body_content = message.Body

        # append the extracted data to the running lists
        all_names.append(message.SenderName)
        # all_dates.append(message.LastModificationTime)
        all_email.append(message.SenderEmailAddress)
        all_body_text.append(message.Body)

extracted_info = pd.DataFrame(columns=["Name", "email", "Date", "Message"])
extracted_info["Name"] = all_names
extracted_info["email"] = all_email
# extracted_info["Date"] = all_dates
extracted_info["Message"] = all_body_text

extracted_info.to_excel("extracted_info.xlsx", index=False)
print(f'{len(extracted_info)} items in the dataframe')
print(
    os.path.join(
        Path.home(), "OneDrive - Prescient", "py", "gitrepo", "extracted_info.xlsx"
    )
)
print("Done extracting info from Outlook emails")
