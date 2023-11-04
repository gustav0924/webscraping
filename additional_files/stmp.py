import smtplib
import os
import ssl
import glob
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd


class Email():
    def __init__(self,sender='sender@email.com',smtp_server='smtp.server.com',port=25, receiver=['receiver1@email.com'],msg = MIMEMultipart(),subject='Testing',body='Hello World\n This is an automated testing message'):
        self.sender = sender
        self.smtp_server = smtp_server
        self.port = port
        self.subject = subject
        self.receiver = receiver        
        self.body = body
        self.smtpObj = smtplib.SMTP(self.smtp_server,self.port)
        self.smtpObj.ehlo()
        #msg initiated
        self.msg = msg
        self.msg['Subject'] = self.subject
        self.msg['From'] = self.sender
        self.msg['To'] = self.receiver[0]
        self.msg['CC'] = ','.join(self.receiver[1:])

    def htmlandimg(self,img_logo,html_file):
        msgImage=MIMEImage(img_logo.read())
        msgImage.add_header('Content-ID', '<Bose_Logo_Black>')
        msgImage.add_header('Content-Disposition','inline')
        self.msg.attach(msgImage)
        html_data = html_file.read()
        html_file.close()
        html_data = html_data.replace('{{Bose_Logo_Black}}', 'cid:Bose_Logo_Black')
        return self.msg.attach(MIMEText(html_data,'html'))

    def textfromxlsx(self,body):
        text=body
        fromxlsx = f'Message from excel: {text}'
        msgtxt = MIMEText(fromxlsx,'plain')
        msgtxt.add_header('Content-Disposition','inline')
        return self.msg.attach(msgtxt)

    def excel(self):
        #### This could be a function, need to check 
        os.chdir(r'\\tj-vault\SupplyChain\...\Spot Buy Reports')
        newdir = os.getcwd()
        file_type = r'\*xlsx'
        files = glob.glob(fr'{newdir}'+file_type)
        file = max(files,key=os.path.getctime)
        excel=open(file,'rb')
        xlsx = MIMEApplication(excel.read())
        xlsx.add_header('Content-Disposition', 'attachment',filename=file[-23:])
        excel.close()
        return self.msg.attach(xlsx)

    def send(self):
        text = self.msg.as_string() if not None else self.body
        with self.smtpObj as server:
            server.sendmail(self.sender,self.receiver,text)
            print('Mail sent')
            server.quit()
    
   


if __name__ == '__main__':
    xlsxmail = pd.read_excel(r"\\tj-vault\SupplyChain\...\Spot_Buy_List.xlsx",sheet_name='Sheet2')
    senders = xlsxmail['From'].dropna().values.tolist()[0]
    receivers = xlsxmail['To'].dropna().values.tolist()
    cC = xlsxmail['CC'].dropna().values.tolist()
    subject = xlsxmail['Subject'].dropna().values.tolist()[0]
    body = xlsxmail['Body'].dropna().values.tolist()[0]
    receivers = receivers + cC
    email = Email()
    email.htmlandimg(open(r'\\tj-vault\SupplyChain\...\Bose_Logo_Black.jpg','rb'),open(r"\\tj-vault\SupplyChain\...\body.html"))
    email.textfromxlsx(body)
    email.excel()
    email.send()