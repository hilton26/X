#!/usr/bin/env python
# coding: utf-8

# # Extract Overdraft Report Attachments from Outlook

print("\n\n################################################")
print("#                                              #")
print("#            START overdrafts.py               #")
print("#                                              #")
print("################################################\n\n")

import os
import time

import win32com.client as win32  # !pip install pywin32

from constants import pthOverdrafts
from utilities import timediff, prior_working_day

start_time = time.time()
print("Connecting to Outlook and locating the latest overdraft report emails ...")

# each source: the sender to match, the subject substrings that must all be
# present, and the suffix used to build the saved file's name
SOURCES = [
    {
        "sender": "pfsautomatedreporting@prescient.co.za",
        "subject_contains": [
            "PRESCIENT INVESTMENT MANAGEMENT Cash Recon",
        ],
        "filename_suffix": "overdrafts_sa.xlsx",
    },
    {
        "sender": "passport_reporting@ntrs.com",
        "subject_contains": ["Report Center - Attachment Notification"],
        "filename_suffix": "_overdrafts_ex-sa.xlsx",
    },
]


def get_sender_smtp(message) -> str:
    """Return the sender's SMTP address, resolving Exchange DNs if needed."""
    try:
        if message.SenderEmailType == "EX":
            return message.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
        return (message.SenderEmailAddress or "").lower()
    except Exception:
        return (message.SenderEmailAddress or "").lower()


outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # newest first

if not os.path.exists(pthOverdrafts):
    raise Exception(f"Folder not found: {pthOverdrafts}")

for source in SOURCES:
    print(f"\n  Searching for the latest email from {source['sender']} ...")

    match = None
    for message in messages:
        try:
            subject = str(message.Subject)
            sender = get_sender_smtp(message)
        except Exception:
            continue
        if sender == source["sender"].lower() and all(
            s in subject for s in source["subject_contains"]
        ):
            match = message
            break  # messages sorted newest-first, so the first hit is the latest

    if match is None:
        print(f"    No matching email found for {source['sender']}.")
        continue

    file_date = match.ReceivedTime.strftime("%Y%m%d")
    pw_day = prior_working_day(match.ReceivedTime).strftime("%Y%m%d")
    save_path = os.path.join(pthOverdrafts, f"{pw_day}_{source['filename_suffix']}")

    saved = False
    for i in range(1, match.Attachments.Count + 1):
        attachment = match.Attachments.Item(i)
        if attachment.FileName.lower().endswith((".xlsx", ".xls")):
            attachment.SaveAsFile(save_path)
            print(f"    Saved: {save_path}")
            saved = True

    if not saved:
        print(
            f"    No spreadsheet attachment found on the latest email from {source['sender']}."
        )

print(f"\n{timediff(start_time, time.time())} extracting overdraft attachments")

print("\n\n################################################")
print("#                                              #")
print("#             END overdrafts.py                #")
print("#                                              #")
print("################################################\n\n")
