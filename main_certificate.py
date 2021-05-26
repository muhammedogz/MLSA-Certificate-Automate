#!/usr/bin/env python3

import os
from certificate import *
from docx import Document


# create output folder if not exist
try:
    os.mkdir("Output")
except OSError:
    pass

filename = "Certificate Template/Event Certificate Template.docx"
doc = Document(filename) 
replace_participant_name(doc, "Muhammed Oğuz Oğuz Oğuz Oğuz Oğuz")
replace_event_name(doc, "WSL + VSCode Edu HD Download")
replace_ambassador_name(doc, "Ayşem Aydoğan")
doc.save('Output/Muhammed Oğuz.docx')

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                print(p.text)