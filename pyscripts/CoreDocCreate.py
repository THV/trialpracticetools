#!/usr/local/bin/Python3.8

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

import sys
import codecs
print(sys.version)
UTF8Writer = codecs.getwriter('utf8')

template = "/Users/tvidal/Desktop/case_questions_template.docx"
#program does not seem to work with a MS Word template
#template = "/Users/tvidal/Desktop/case_questions_template.dotx"

document = MailMerge(template)
print(document.get_merge_fields())

# document.merge({
#    "case_name": "Keller v. Steep Hill",
#    "court": 'San Francisco Superior Court',
#    "c/m": '21175.00005 / Jmîchaeĺe Keller shareholder litigation',
#    "ct_file_no": 'GCC-123'
#    })

document.merge(
    case_name='Keller v. Steep Hill',
    court='San Francisco Superior Court',
    cm_no='21175.00005',
    ct_file_no='GCC-123-456')

document.write('/Users/tvidal/Desktop/questions_probs.docx')
