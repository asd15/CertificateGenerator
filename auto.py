import pandas as pd
import smtplib

from PIL import Image, ImageDraw, ImageFont

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from twilio.rest import Client

account_sid = "" ##twilo accound sid
auth_token = "" ##twilo auth token
client = Client(account_sid, auth_token)

frmaddr = #your email id
pw = #your password

def takefile(filename):
    print(filename)
    xlfile = pd.read_excel(filename, 'Sheet1')

    for i in xlfile.index:
        image = Image.open('certificate.PNG')
        draw = ImageDraw.Draw(image)
        name = xlfile['Name'][i]
        date = xlfile['Date'][i]
        email = xlfile['Email'][i]
        phone = xlfile['Phone'][i]
        phone2 = str(phone)
        plus = "+91"
        phonenum = plus + phone2
        print(name)
        font = ImageFont.truetype('LittleLordFontleroyNF.ttf', size=200)
        color = 'rgb(0, 0, 0)'
        draw.text((100, 470), name, fill=color, font=font)
        font2 = ImageFont.truetype('Arial.ttf', size=30)
        draw.text((130, 850), date, fill=color, font=font2)
        imageName = name + ".PNG"
        image.save(imageName)

        toaddr = xlfile['Email'][i]
        msg = MIMEMultipart()

        msg['From'] = frmaddr
        msg['To'] = toaddr
        msg['Subject'] = "Certificate for Completion of Python Course"

        body = '''Thank you for enrolling in our course. We hope you made the best out of it. We will come with more courses soon. Stay
        Tuned. Your certificated is attached with this mail.

        Thanks and Regards 
        Team learnovate.org
        '''
        msg.attach(MIMEText(body, 'plain'))

        filename = name + ".PNG"
        attachment = open(imageName, "rb")
        part = MIMEBase('application', 'octet-stream')

        part.set_payload((attachment).read())

        encoders.encode_base64(part)

        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)

        server.starttls()
        server.login(frmaddr, pw)

        text = msg.as_string()

        server.sendmail(frmaddr, toaddr, text)

        server.quit()

        message = client.messages.create(
                              from_='whatsapp:+14155238886',
                              body='Thank you for attending our course in Python in Data Science. Your certificate has been mailed to you on your email id: '+email,
                              to='whatsapp:'+phonenum
                          )
        print(message.sid)
