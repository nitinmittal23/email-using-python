import smtplib
import xlrd 
import time
def send_email(subject, msg, email):
        server = smtplib.SMTP('smtp.gmail.com:587')
        server.ehlo()
        server.starttls()
        server.login('senderEmail', 'password')
        message = 'Subject: {} \n\n{}'.format(subject, msg)
        server.sendmail('senderEmail', 'receiverEmail',message)
        server.quit()
        
        
loc = ("LocationOfExcelSheet") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
  
for i in range(0,sheet.nrows): 
    subject = 'open Source Development Workshop'
    msg = 'Hey {}, \nThis is to inform you that your registration is confirmed for Open source workshop in MAIT in association with Student Chapter of Coding Blocks. As we have Limited seats, the entry will be based on First Come Basis and no entry will be given after the room gets FILLED. So, be there by 10:30 AM to get your seat. \nVenue: Lab 131-132, CSE Department, MAIT \nThanks And Regards \nCampus Blocks\n\n\n'.format(sheet.cell_value(i, 1))
    email = '{}'.format(sheet.cell_value(i, 2))
    send_email(subject,msg, email)
    time.sleep(2) 