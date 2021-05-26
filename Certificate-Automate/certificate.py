#!/usr/bin/env python3
import re
from docx import Document

def docx_replace_regex(doc_obj, regex , replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            print(p.text)
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
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

filename = "Certificate-Automate/Event Certificate Template.docx"
doc = Document(filename)

replace_participant_name(doc, "Muhammed Oğuz")

doc.save('Output/result1.docx')