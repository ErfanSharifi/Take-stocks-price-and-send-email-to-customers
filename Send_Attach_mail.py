# -*- coding: utf-8 -*-
#!/usr/bin/python

import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 


class Send():
    
    def __init__(self):

        pass

    def GiveMail(self, mail):

        fromaddr = 'sharifi.stockreport@gmail.com'
        toaddr = mail
        
        # instance of MIMEMultipart 
        msg = MIMEMultipart() 
        
        # storing the senders email address   
        msg['From'] = fromaddr 
        
        # storing the receivers email address  
        msg['To'] = toaddr 
        
        # storing the subject  
        msg['Subject'] = """اطلاعات ماهانه پورتفوی شما"""
        
        # string to store the body of the mail 
        body =  """
                                             .سلام و روز بخیر
        .اطلاعات پورتفوی شما در این ماه خدمتتان ارسال گردید
                                               .پر سود باشید"""
        
        # attach the body with the msg instance 
        msg.attach(MIMEText(body, 'plain')) 
        
        # open the file to be sent  
        path = "C:/Users/Erfan/OneDrive/Projects/SSM_Test/Outputs/99.6.31.xlsx"
        filename = "99.6.31.xlsx"
        attachment = open(path, "rb") 
        
        # instance of MIMEBase and named as p 
        p = MIMEBase('application', 'octet-stream')
        
        # To change the payload into encoded form 
        p.set_payload((attachment).read()) 
        
        # encode into base64 
        encoders.encode_base64(p) 
        
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
        
        # attach the instance 'p' to instance 'msg' 
        msg.attach(p) 
        
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        
        # start TLS for security 
        s.starttls() 
        
        # Authentication 
        s.login(fromaddr, "Erfan123@#$%") 
        
        # Converts the Multipart msg into a string 
        text = msg.as_string() 
        
        # sending the mail 
        s.sendmail(fromaddr, toaddr, text) 
        
        # terminating the session 
        s.quit() 

# def main():
    
#     mail = 'e.sharify@gmail.com'
#     obj = Send()
#     obj.GiveMail(mail)
# main()