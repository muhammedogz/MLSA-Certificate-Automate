#!/usr/bin/env python3
import re
import os
from docx import Document


# https://stackoverflow.com/a/42829667/11970836
# This function replace data and keeps style
def docx_replace_regex(doc_obj, regex , replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                print(inline[i].text)
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex , replace)


def replace_participant_name(doc, name):
    string = "Samplefirstname Samplelastname"
    reg = re.compile(r""+string)
    replace = r""+name
    docx_replace_regex(doc, reg , replace)

def replace_event_name(doc, event):
    string = "{INSERT EVENT NAME}"
    reg = re.compile(r""+string)
    replace = r""+event
    docx_replace_regex(doc, reg , replace)

def replace_ambassador_name(doc, name):
    string = "{student ambassador name}"
    reg = re.compile(r""+string)
    replace = r""+name
    docx_replace_regex(doc, reg , replace)

# create output folder if not exist
try:
    os.mkdir("Output")
except OSError:
    pass

