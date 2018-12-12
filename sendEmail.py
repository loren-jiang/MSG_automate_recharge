import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

def send_mail(send_from,send_to,subject,text,files,server,port,username='',password='',isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(files, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="'+files+'"')
    msg.attach(part)

    #context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
    #SSL connection only working on Python 3+
    smtp = smtplib.SMTP(server, port)
    if isTls:
        smtp.ehlo()
        smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

TO = 'loren.jiang@gmail.com'
SUBJECT = 'Python Email'
TEXT = 'Here is the message'

gmail_sender = 'loren.jiang@gmail.com'
gmail_passwd = 'Kuni8kuni'

server = 'smtp.gmail.com'
port = 587
file = 'TEST.xlsx'
send_mail(TO,gmail_sender,SUBJECT,TEXT,file,server,port,gmail_sender,gmail_passwd,isTls=True)