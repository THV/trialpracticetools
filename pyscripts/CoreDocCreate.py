#!/usr/local/bin/Python3.8

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

import sys
import codecs
UTF8Writer = codecs.getwriter('utf8')

#Associate the Core Document Templates
coredoc01_tmpl = "/Users/tvidal/Dropbox/Templates/99.Core Docs/CoreDoc01_finder_template.docx"
coredoc02_tmpl = "/Users/tvidal/Dropbox/Templates/99.Core Docs/CoreDoc02_litplan_template.docx"
coredoc10_tmpl = "/Users/tvidal/Dropbox/Templates/99.Core Docs/CoreDoc10_case_questions_template.docx"

#Assign the templates to the core merge documents
coredoc01 = MailMerge(coredoc01_tmpl)
coredoc02 = MailMerge(coredoc02_tmpl)
coredoc10 = MailMerge(coredoc10_tmpl)

#Assign kwargs
kwargs = ({
    "case_name": 'Keller v. Steep Hill',
    "court": 'San Francisco Superior Court',
    "c/m#": '21175.00005 / Jmîchaeĺe Keller shareholder litigation',
    "ct_file_no": 'GCC-123'
    })

#Merge the documents
coredoc01.merge(**kwargs)
coredoc02.merge(**kwargs)
coredoc10.merge(**kwargs)

#Write the documents
coredoc01.write('/Users/tvidal/Desktop/Core Doc 01 - Info Sheet.docx')
coredoc02.write('/Users/tvidal/Desktop/Core Doc 02 - Litigation Plan.docx')
coredoc10.write('/Users/tvidal/Desktop/Core Doc 10 - Questions & Probs.docx')
