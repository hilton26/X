#!/usr/bin/env python
# coding: utf-8

# # **Extract text, links, images, tables from Pdf with Python | PyMuPDF, PyPdf, PdfPlumber tutorial**
#
# https://www.youtube.com/watch?v=G0PApj7YPBo
#
# https://pypdf.readthedocs.io/en/stable/
#
# https://www.ibm.com/docs/en/db2-event-store/2.0.0?topic=notebooks-markdown-jupyter-cheatsheet

# libraries, libraries!
import time

# !pip install  pypdf
# !pip install python-docx
from datetime import datetime
from pypdf import PdfReader
from docx import Document  # https://python-docx.readthedocs.io/en/latest/
from tqdm import tqdm
import os
import pandas as pd
from utilities import timediff
from constants import pthPy, pthTest

# set path to the pdf to be parsed
start_time = time.time()

print("Saving the pdf text to Word ...", "\n")

# get the pdf file path and name
df = pd.read_excel(pthPy, sheet_name="arc", usecols="AQ", nrows=1)
df
path_pdf = df.iloc[0, 0]
print(path_pdf)

# apply the pdf reader to the pdf
reader = PdfReader(path_pdf)

# print stats on number of pdf pages
lx = len(reader.pages)
print("", f"{lx} page{'' if lx == 1 else 's'} in the pdf")

# print name of the eventual Word doc
# https://www.w3schools.com/python/ref_string_rfind.asp#:~:text=The%20rfind()%20method%20finds,as%20the%20rindex()%20method.
doc_name = path_pdf[path_pdf.rfind("\\") + 1 : path_pdf.rfind(".")] + ".docx"
print("", os.path.join(pthTest, f"{doc_name}"), "\n")

# create a Word document to hold the extracted pdf text https://www.youtube.com/watch?v=re-rBqvDUPM
doc = Document()

# get text from every page in the pdf and add it to the Word doc
for page in tqdm(range(len(reader.pages))):
    pg = reader.pages[page]
    text = pg.extract_text()
    # print(text)
    doc.add_paragraph(text)
    # doc.add_page_break()

doc.save(os.path.join(pthTest, f"{doc_name}"))  # save in python exports folder
doc.save(
    path_pdf[0 : path_pdf.rfind("\\") + 1] + f"{doc_name}"
)  # save in the local pdf folder

print(f"{timediff(start_time, time.time())} saving the pdf text to Word")
print(path_pdf[0 : path_pdf.rfind("\\") + 1] + f"{doc_name}")
