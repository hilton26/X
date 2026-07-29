# !pip install pywin32
import os
import win32com.client
from datetime import datetime

# =========================
# CONFIGURATION
# =========================
TARGET_SUBJECT = "PRESCIENT INVESTMENT MANAGEMENT Cash Recon"

SAVE_FOLDER = r"\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Overdrafts"

# Ensure folder exists
if not os.path.exists(SAVE_FOLDER):
    raise Exception(f"Folder not found: {SAVE_FOLDER}")

# =========================
# CONNECT TO OUTLOOK
# =========================
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items

# Sort by newest first (recommended for performance)
messages.Sort("[ReceivedTime]", True)

processed_count = 0

print("Starting extraction...")

# =========================
# LOOP THROUGH EMAILS
# =========================
for message in messages:
    try:
        subject = str(message.Subject)

        if TARGET_SUBJECT in subject:
            
            received_time = message.ReceivedTime
            file_date = received_time.strftime("%Y%m%d")

            # Process attachments
            attachments = message.Attachments

            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                filename = attachment.FileName.lower()

                # Only Excel files
                if filename.endswith(".xlsx") or filename.endswith(".xls"):
                    
                    new_filename = f"{file_date}_PIM_Cash_Recon.xlsx"
                    save_path = os.path.join(SAVE_FOLDER, new_filename)

                    # Save (this will overwrite if file exists)
                    attachment.SaveAsFile(save_path)

                    print(f"Saved: {save_path}")
                    processed_count += 1

    except Exception as e:
        print(f"Error processing email: {e}")

print(f"\nCompleted. Total files saved: {processed_count}")