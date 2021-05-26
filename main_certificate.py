#!/usr/bin/env python3

import os
from certificate import *
from docx import Document
from docx2pdf import convert


# create output folder if not exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass

def create_docx_files(filename, list_participate, event, ambassador):


    for participate in list_participate:
        # use original file everytime
        doc = Document(filename)

        replace_participant_name(doc, participate)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)
        doc.save('Output/Doc/{}.docx'.format(participate))
        convert('Output/Doc/{}.docx'.format(participate), 'Output/Pdf/{}.pdf'.format(participate))
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        print(p.text)

    

filename = "Certificate Template/Event Certificate Template.docx"

list_participate = ["Muhammed Oğuz", "Ayşegül Aydoğan", "Joe Biden", "Barack Obama"]

create_docx_files(filename, list_participate, "WSL and VSCode ha", "Aysem Yas")



