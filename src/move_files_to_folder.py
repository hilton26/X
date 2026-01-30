#!/usr/bin/env python
# coding: utf-8

# Gemini query 26 Nov 2025: "python script to move files with a given modified date to a different named folder"

from utilities import timediff, move_files_by_modified_date
from datetime import date

# if __name__ == "__main__":
# Define your source and destination directories
source_folder = r"\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Derivative Cover"
destination_folder = r"\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Derivative Cover\2024_and_prior"

# Define the target modification date (e.g., files older than or equal to 2025-06-30)
# Year, Month, Day
target_mod_date = date(2025, 6, 30)

move_files_by_modified_date(source_folder, destination_folder, target_mod_date)
