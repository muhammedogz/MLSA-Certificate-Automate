#!/usr/bin/env python3

import os
from certificate import *
from docx import Document


# create output folder if not exist
try:
    os.mkdir("Output")
except OSError:
    pass

def create_docx_files(doc, dict_participant, event, ambassador):
    replace_participant_name(doc, dict_participant)
    replace_event_name(doc, event)
    replace_ambassador_name(doc, ambassador)
    doc.save('Output/{}.docx'.format(dict_participant))

filename = "Certificate Template/Event Certificate Template.docx"
doc = Document(filename) 

create_docx_files(doc, "Domates Patates", "WSL and VSCode ha", "Aysem Yas")

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                print(p.text)


