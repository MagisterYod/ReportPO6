import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


smtp_server = '172.23.100.20'
smtp_port = 25
sender_email = 'asd_water@polymetal.ru'
receiver_email = [
    # 'OS2@polymetal.ru',
    'kuznecovia@polymetal.ru'
]


with smtplib.SMTP(smtp_server, smtp_port) as server:
    body = f'Не добрый день!\n\nОшибка: {1}\n\nФайл: {1}\n\n'
    message = MIMEMultipart()
    message["Subject"] = f'{datetime.datetime.now().strftime("%d.%m.%Y")} {1}'
    message["From"] = sender_email
    message["To"] = receiver_email
    message_text = MIMEText(body.encode('utf-8'), 'plain', 'utf-8')
    message.attach(message_text)
    server.ehlo()  # Может быть опущено
    server.sendmail(sender_email, receiver_email, message.as_string())
    server.quit()
