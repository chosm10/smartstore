import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import os
from mimetypes import guess_type

# to를 리스트로 주면 여러명에게 발송 가능
def sendmail(to, subject, body, attachs):
    smtp_server = smtplib.SMTP('10.100.32.29', 25)#smtplib.SMTP('130.102.102.70', 25)# 웹메일  #G메일 smtplib.SMTP('smtp.gmail.com', 587)#  #
    smtp_server.ehlo()
    smtp_server.starttls()

    msg = MIMEMultipart('mixed')
    msg.set_charset('utf-8')
    From = 'rdep001@hyundai-ite.com'
        
    msg['From'] = From
    msg['To'] = ",".join(to)
    msg['Subject'] = subject

    bodyPart = MIMEText(body, 'html', 'utf-8')
    msg.attach(bodyPart)

    # 첨부파일 경로가 존재하면 첨부
    if attachs:
        for attach in attachs:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(open(attach, 'rb').read())
            encoders.encode_base64(part)
            filename = os.path.basename(attach)
            part.add_header('Content-Disposition', 'attachment', filename=filename)#('UTF-8', '', filename)
            
            msg.attach(part)

    smtp_server.sendmail(From, to, msg.as_string())
    smtp_server.quit()