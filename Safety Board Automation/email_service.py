import win32com.client as win32
import sys
import os

def email_sender():
    """This funcitons sends a email"""

    # Creates a string of recievers from Emails.txt
    emails_txt = os.path.join(os.path.dirname(sys.path[0]), "Input/Emails.txt").replace("\\", "/")
    with open(emails_txt) as f:
        emails = f.readlines()
    str_recievers = ''
    for email in emails:
        email = email.replace('\n', '')
        str_recievers += f"{email}; "
    str_recievers = str_recievers[:-2]

    # Set up the email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = str_recievers
    mail.Subject = 'Safety Board Analytics'
    mail.Body = ''
    # mail.HTMLBody = '' #this field is optional

    # Set up the attachments
    attachment1  = (os.path.join(os.path.dirname(sys.path[0]), "Output/Inspections Data Mining AO Analytics.xlsx")).replace("\\", "/")
    attachment2  = (os.path.join(os.path.dirname(sys.path[0]), "Output/Question level AO Analytics.xlsx")).replace("\\", "/")
    attachment3  = (os.path.join(os.path.dirname(sys.path[0]), "Output/Inspections Data Mining FSI Analytics.xlsx")).replace("\\", "/")
    attachment4  = (os.path.join(os.path.dirname(sys.path[0]), "Output/Question level FSI Analytics.xlsx")).replace("\\", "/")
    
    mail.Attachments.Add(attachment1)
    mail.Attachments.Add(attachment2)
    mail.Attachments.Add(attachment3)
    mail.Attachments.Add(attachment4)

    # Send the email
    mail.Send()