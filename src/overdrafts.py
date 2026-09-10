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
        "sender": "pfsautomatedreporting@prescient.co.za",
        "subject_contains": [
            "Unconfirmed Cash Balances - PRESCIENT INVESTMENT MANAGEMENT",
        ],
        "filename_suffix": "unconfirmed_cash_balances_sa.xlsx",
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


SEARCH_TIMEOUT_SECONDS = 180  # give up on a source if nothing matches within 3 minutes
POLL_INTERVAL_SECONDS = 15

outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

if not os.path.exists(pthOverdrafts):
    raise Exception(f"Folder not found: {pthOverdrafts}")

for source in SOURCES:
    print(f"\n  Searching for the latest email from {source['sender']} ...")

    match = None
    search_start = time.time()
    while match is None:
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # newest first
        for message in messages:
            try:
                subject = str(message.Subject)
                sender = get_sender_smtp(message)
            except Exception:
                continue
            if not all(s in subject for s in source["subject_contains"]):
                continue

            is_direct_match = sender == source["sender"].lower()
            is_forwarded_match = False
            if not is_direct_match:
                # the email may have been forwarded by someone else, so also check
                # whether the original sender's address appears in the message body
                try:
                    body = str(message.Body or "")
                except Exception:
                    body = ""
                is_forwarded_match = source["sender"].lower() in body.lower()

            if is_direct_match or is_forwarded_match:
                match = message
                break  # messages sorted newest-first, so the first hit is the latest

        if match is None:
            if time.time() - search_start >= SEARCH_TIMEOUT_SECONDS:
                break
            time.sleep(POLL_INTERVAL_SECONDS)

    if match is None:
        print(
            f"    No matching email found for {source['sender']} "
            f"after {SEARCH_TIMEOUT_SECONDS // 60} minutes."
        )
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
