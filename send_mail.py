def get_credentials():
    with open('credentials.txt') as f:
        # файл, в первой строчке которого находится логин, а во второй пароль
        login = f.readline().strip()
        password = f.readline().strip()
    return login, password


def send(text, subject, from_email, to_email, host='smtp.gmail.com'):
    # Import smtplib for the actual sending function
    import smtplib

    # Import the email modules we'll need
    from email.mime.text import MIMEText

    # # Open a plain text file for reading.  For this example, assume that
    # # the text file contains only ASCII characters.
    # fp = open(textfile, 'rb')
    # # Create a text/plain message
    # msg = MIMEText(fp.read())
    # fp.close()
    msg = MIMEText(text)

    # me == the sender's email address
    # you == the recipient's email address
    msg['Subject'] = subject # 'The contents of %s' % textfile
    msg['From'] = from_email #me
    msg['To'] = to_email #you

    # Send the message via our own SMTP server, but don't include the
    # envelope header.
    s = smtplib.SMTP(host)
    s.starttls()
    s.login(*get_credentials())
    s.sendmail(from_email, [to_email], msg.as_string())
    s.quit()


# =============
import smtplib
import os
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


def send_mail(send_from, send_to, subject, text, files=None,
              server="127.0.0.1"):
    #print(send_to)
    assert isinstance(send_to, list)

    msg = MIMEMultipart(
        From=send_from,
        To=COMMASPACE.join(send_to),
        Date=formatdate(localtime=True),
        Subject=subject
    )
    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            msg.attach(MIMEApplication(
                fil.read(),
                Content_Disposition='attachment; filename="%s"' % basename(f),
                Name=basename(f)
            ))

    smtp = smtplib.SMTP(server)
    smtp.starttls()
    #print(server)
    smtp.login(*get_credentials())
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()



#text = 'Hi! My dear friend!'
#subject = 'Test.Test'
#from_email, to_email = 'n.k.kruglyak@gmail.com', 'o.a.yunin@gmail.com '
#send_to = [to_email]
#send(text, subject, from_email, to_email)
#name_file = 'Кругляк.xls'
#files = os.path.join(os.path.abspath(os.path.dirname(__file__)), name_file)
#send_mail(from_email,send_to, subject, text, files=[files],server='smtp.gmail.com')