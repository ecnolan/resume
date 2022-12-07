"""
    File:        resume.py
    Author:      Eva Nolan
    Description: Convert csv resume file into pdf and html files
"""

import pandas as panda
import numpy as np
import aspose.words as aw
import csv

# Import csv
def create_fonts(builder):
    body = builder.font
    body.size = 11
    body.bold = False
    body.name = "Garamond"

    header = body
    header.bold = True

    return (body,header)


def write_resume(resume_csv, word_doc, html_doc):
    data = panda.read_csv(resume_csv)
    print("1")
    word = aw.Document()
    print("2")

    # create a document builder object
    builder = aw.DocumentBuilder(word)
    print("3")

    # add text to the document
    builder.write("Hello world!")
    print("4")

    builder = aw.DocumentBuilder(word)
    # create fonts
    (body, header) = create_fonts(builder)

    # set paragraph formatting
    paragraphFormat = builder.paragraph_format
    paragraphFormat.first_line_indent = 8
    paragraphFormat.alignment = aw.ParagraphAlignment.JUSTIFY
    paragraphFormat.keep_together = True

    # add text
    builder.writeln("A whole paragraph.")

    # save document
    word.save(word_doc)
    print("5")


if __name__ == "__main__":
    resume_csv = "resume.csv"
    word_doc = "resume_summer_2022.docx"
    html_doc = "resume_summer_2022.txt"
    print("hello")
    write_resume(resume_csv, word_doc, html_doc)
    # test() # use for testing, comment when done
