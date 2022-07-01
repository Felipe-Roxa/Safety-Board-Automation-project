import win32com.client as win32

def email_sender():
    """This funcitons sends a email"""
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'felsique@amazon.com'
    mail.Subject = 'Safety Board Analytics'
    mail.Body = 'Message body'
    mail.HTMLBody = '' #this field is optional

    # To attach a file to the email (optional):
    attachment1  = "C:/Users/felsique/Desktop/Safety Board/Output/Inspections Data Mining AO Analytics.xlsx"
    attachment2  = "C:/Users/felsique/Desktop/Safety Board/Output/Question level AO Analytics.xlsx"
    attachment3  = "C:/Users/felsique/Desktop/Safety Board/Output/Inspections Data Mining FSI Analytics.xlsx"
    attachment4  = "C:/Users/felsique/Desktop/Safety Board/Output/Question level FSI Analytics.xlsx"
    mail.Attachments.Add(attachment1)
    mail.Attachments.Add(attachment2)
    mail.Attachments.Add(attachment3)
    mail.Attachments.Add(attachment4)

    mail.Send()