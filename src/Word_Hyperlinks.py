#!/usr/bin/env python
# coding: utf-8

# ### Python extracts links (hyperlink) and text from word/docx
# https://stdworkflow.com/441/python-extracts-links-hyperlink-and-text-from-word-docx

# How to index into a dictionary?<br>
# https://stackoverflow.com/questions/4326658/how-to-index-into-a-dictionary
# 
# Python typeerror: ‘int’ object is not iterable Solution<br>
# https://careerkarma.com/blog/python-typeerror-int-object-is-not-iterable/#:~:text=TypeErrors%20are%20a%20common%20type,iterable%20rather%20than%20a%20number.

# In[71]:


### Libraries ###
import zipfile
import re
import json
import base64
from docx import Document
from os.path import basename
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from bs4 import BeautifulSoup


### Constants ###
pth = 'J:\\PProfile Operations\\Legal\\Legal and Compliance Framework\\Hilton\\Client Mandates\\'


filename = pth + 'POIF Rules.docx'

### Functions ###
def get_linked_text(soup):

    links = []

    # This kind of link has a corresponding URL in the _rel file.
    for tag in soup.find_all("hyperlink"):
        # try/except because some hyperlinks have no id.
        try:
            links.append({"id": tag["r:id"], "text": tag.text})
        except:
            pass

    # This kind does not.
    for tag in soup.find_all("instrText"):
        # They're identified by the word HYPERLINK
        if "HYPERLINK" in tag.text:
            # Get the URL. Probably.
            url = tag.text.split('"')[1]

            # The actual linked text is stored nearby tags.
            # Loop through the siblings starting here.
            temp = tag.parent.next_sibling
            text = ""

            while temp is not None:
                # Text comes in <t> tags.
                maybe_text = temp.find("t")
                if maybe_text is not None:
                    # Ones that have text in them.
                    if maybe_text.text.strip() != "":
                        text += maybe_text.text.strip()

                # Links end with <w:fldChar w:fldCharType="end" />.
                maybe_end = temp.find("fldChar[w:fldCharType]")
                if maybe_end is not None:
                    if maybe_end["w:fldCharType"] == "end":
                        break

                temp = temp.next_sibling

            links.append({"id": None, "href": url, "text": text})

    return links

#if __name__ == '__main__':
    #file_name = filename
    #archive = zipfile.ZipFile(file_name, "r")
    #file_data = archive.read("word/document.xml")
    #doc_soup = BeautifulSoup(file_data, "xml")
    #linked_text = get_linked_text(doc_soup)
    #print(linked_text)
    
if __name__ == '__main__':
   linked_text = get_linked_text(BeautifulSoup(zipfile.ZipFile(filename, "r").read("word/document.xml"), "xml"))

#print(linked_text)

def links(filename):
    return get_linked_text(BeautifulSoup(zipfile.ZipFile(filename, "r").read("word/document.xml"), "xml"))  
    #print(get_linked_text(BeautifulSoup(zipfile.ZipFile(filename, "r").read("word/document.xml"), "xml")))
    
def get_hyperlinks(text,filename):
    linked_text = links(filename)
    hlink = []
    for i in range(len(linked_text)):
        m = linked_text[i][list(linked_text[i])[1]]
    if text in m:
        hlink.append(m)
    return hlink


# In[72]:


get_hyperlinks('Segregated Clients',pth + 'POIF Rules.docx')


# In[ ]:




