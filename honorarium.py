
from datetime import date, datetime
from docx import Document
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from io import StringIO
from gmail_interaction import create_message, create_draft

import docx2txt
import os
import shutil

#First, we get the specifics of this review
author          = input("Author (full name): ")
title           = input("Title: ")
honorarium      = input("Honorarium (number value): ")
due             = input("Due date (mm/dd/yyyy): ")
reviewer        = input("Reviewer last name: ")
#reviewer_address = input("Reviewer email address: ")

# We also need the day's date Woudl be nice if it could catch both cases
today = date.today() # dd/mm/YY
d1 = today.strftime("%m/%d/%Y")

#to make the date look nice for the email
date_object = datetime.strptime(due, '%m/%d/%Y')
full_date = date_object.strftime('%B %d, %Y')

#Note: maybe make this a dialogue box or automated one day; it would make the initial questions easier to fill out 
## because less clicking between tabs !Chrome extension!

#Now, we format the information to make it palatable for the docxptl library in the form of a.... dictionary!
def get_context(): 
    return{'author': author,
        'title': title,
        'due': due,
        'honorarium': honorarium,
        'honorarium2': int(honorarium)*2,
        'today': d1,
          'full_date': full_date,
          'reviewer' : reviewer
          }

#We then put this information into the logsheet and email templates
def make_logsheet():
    logsheet_title = reviewer + " honorarium logsheet.docx"

    template = DocxTemplate("honorarium_logsheet_template.docx") #Insert the name of the template file here

    context = get_context()
    target_file = StringIO()
    template.render(context)
    template.save(logsheet_title) 
    
    #Saves to the honorarium logsheet folder
    filepath = os.path.join(r"/Users/karenli/Box/Departmental (rutgerspress2)/Acquisitions/Interns--Incl. Comp Copy Orders/Interns--Incl. Comp Copy Orders/Peter's interns/Honorarium Logsheets & W9s/2021 Honorarium Logsheets",logsheet_title)
    shutil.move(logsheet_title, filepath)

def write_email():
    template = DocxTemplate("honorarium_email_template.docx") #Insert the name of the template file here

    context = get_context()
    target_file = StringIO()
    template.render(context)
    template.save("temp.docx")

    return docx2txt.process("temp.docx")

def draft_email():
    subject = title      #TODO: Parse main tile base on colon
    email_body = write_email()
    start_emailer(recipient, subject, email_body)
    

    
make_logsheet()
write_email()

#draft_email
print("Sent to " + reviewer + ", due " + due) #for my record
print("Review(" + author + "): sent to " + reviewer) #for Peter's spreadsheet
#woudl be nice to have these automatically append