import win32com.client as win32

def email_sender():
    """This funcitons sends a email"""
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'felsique@amazon.com'
    mail.Subject = 'Message subject'
    mail.Body = 'Message body'
    mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment1  = "C:/Users/felsique/Desktop/Safety Board/Output/safety_board_ao_inspection_data_mining.xlsx"
    attachment2  = "C:/Users/felsique/Desktop/Safety Board/Output/safety_board_ao_question_level.xlsx"
    attachment3  = "C:/Users/felsique/Desktop/Safety Board/Output/safety_board_fsi_inspections_data_mining.xlsx"
    attachment4  = "C:/Users/felsique/Desktop/Safety Board/Output/safety_board_fsi_question_level.xlsx"
    mail.Attachments.Add(attachment1)
    mail.Attachments.Add(attachment2)
    mail.Attachments.Add(attachment3)
    mail.Attachments.Add(attachment4)

    mail.Send()