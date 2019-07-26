# -*- coding: utf-8 -*-
"""
@author: Maciej Ulaszewski
github: github.com/ulaszewskim
"""

import os
import sys
import comtypes.client


def word_to_pdf(directory):
    """Convert all doc/docx file in directory to PDF file"""
    files = []
    files_list = os.listdir(directory)
    for file in files_list:
        if file.lower().endswith(('.docx', '.doc')):
            files.append(file)
    del files_list

    for file in files:
        doc_file = os.path.join(directory, file)
        pdf_file = os.path.join(directory, os.path.splitext(file)[0]+'.pdf')
        try:
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(doc_file)
            doc.SaveAs(pdf_file, FileFormat=17)
            doc.Close()
            word.Quit()
            print('Created: {}'.format(os.path.basename(file)))
        except:
            print('Error with file: {}:'.format(file))
            print('    {}'.format(sys.exc_info()[1]))
