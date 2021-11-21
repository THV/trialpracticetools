#!/usr/local/bin/Python

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

import sys
import os # needed for making the output directory
import codecs
UTF8Writer = codecs.getwriter('utf8')

#Get Case Information
our_case = input("Enter our case name: ")
case_no = input("Enter our case/matter Number: ")
court_name = input("Enter the court or other tribunal: ")
court_file_no = input("Enter the court file number: ")

#Create Template Folder and Output Folder
template_folder = "/Users/tvidal/Dropbox/Templates/"
output_folder = "/Users/tvidal/Desktop/Core docs - " + our_case
access_rights = 0o755 #access rights for output folder
os.mkdir(output_folder, access_rights)

#Associate the Core Document Templates
coredoc01_tmpl = template_folder + "10.Finder/10.00.Case Info Sheet.docx"
coredoc02_tmpl = template_folder + "01.The Plan/01.00.THV Litigation Checklist.docx"
coredoc03_tmpl = template_folder + "02.Summary/02.00.One-Sheet Case Summary.docx"
coredoc04_tmpl = template_folder + "02.Summary/02.01.Cast of Characters.docx"
coredoc05_tmpl = template_folder + "02.Summary/02.02.Case Chronology.docx"
coredoc06_tmpl = template_folder + "10.Finder/10.01.Witness List.docx"
coredoc07_tmpl = template_folder + "09.Exhibit List/09.00.Exhibit List.docx"
coredoc08_tmpl = template_folder + "16.Law-Trial Memo/16.01.Issues List-Prima Facie Elements.docx"
coredoc09_tmpl = template_folder + "/13.Damages List/13.01.Damages List.docx"
coredoc10_tmpl = template_folder + "/01.The Plan/01.01.Questions-Problems.docx"
coredoc12_tmpl = template_folder + "/01.The Plan/01.20.Discovery Plan.docx"
coredoc13_tmpl = template_folder + "/01.The Plan/01.21.Discovery Tracking Record.docx"

#Assign the templates to the core merge documents
coredoc01 = MailMerge(coredoc01_tmpl)
coredoc02 = MailMerge(coredoc02_tmpl)
coredoc03 = MailMerge(coredoc03_tmpl)
coredoc04 = MailMerge(coredoc04_tmpl)
coredoc05 = MailMerge(coredoc05_tmpl)
coredoc06 = MailMerge(coredoc06_tmpl)
coredoc07 = MailMerge(coredoc07_tmpl)
coredoc08 = MailMerge(coredoc08_tmpl)
coredoc09 = MailMerge(coredoc09_tmpl)
coredoc10 = MailMerge(coredoc10_tmpl)
coredoc12 = MailMerge(coredoc12_tmpl)
coredoc13 = MailMerge(coredoc13_tmpl)

#Assign kwargs
kwargs = ({
    "case_name": our_case,
    "court": court_name,
    "c/m#": case_no,
    "ct_file_no": court_file_no
    })

#Merge the documents
coredoc01.merge(**kwargs)
coredoc02.merge(**kwargs)
coredoc03.merge(**kwargs)
coredoc04.merge(**kwargs)
coredoc05.merge(**kwargs)
coredoc06.merge(**kwargs)
coredoc07.merge(**kwargs)
coredoc08.merge(**kwargs)
coredoc09.merge(**kwargs)
coredoc10.merge(**kwargs)
coredoc12.merge(**kwargs)
coredoc13.merge(**kwargs)

#Write the documents
filename_end =  "- " + our_case + ".docx"


coredoc01.write(output_folder + '/Core Doc 01 - Info Sheet' + filename_end)
coredoc02.write(output_folder + '/Core Doc 02 - Litigation Plan' + filename_end)
coredoc03.write(output_folder + '/Core Doc 03 - One Sheet Case Summary' + filename_end)
coredoc04.write(output_folder + '/Core Doc 04 - Cast of characters' + filename_end)
coredoc05.write(output_folder + '/Core Doc 05 - Case chronology' + filename_end)
coredoc06.write(output_folder + '/Core Doc 06 - Witness List' + filename_end)
coredoc07.write(output_folder + '/Core Doc 07 - Exhibits list' + filename_end)
coredoc08.write(output_folder + '/Core Doc 08 - Issues list' + filename_end)
coredoc09.write(output_folder + '/Core Doc 09 - Damages List-Summary' + filename_end)
coredoc10.write(output_folder + '/Core Doc 10 - Questions & Probs' + filename_end)
coredoc12.write(output_folder + '/Core Doc 12 - Discovery Plan' + filename_end)
coredoc13.write(output_folder + '/Core Doc 13 - Discovery Tracking System' + filename_end)

