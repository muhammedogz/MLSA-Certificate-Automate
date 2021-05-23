#!/usr/bin/env python3

from docx import Document


print("hello world")


filename = "Certificate-Automate/Event Certificate Template.docx"
f = open(filename, 'rb')
doc = Document(f)

for section in doc.sections:
    header = section.header
    for paragraph in header.paragraphs:
        print(paragraph.text)