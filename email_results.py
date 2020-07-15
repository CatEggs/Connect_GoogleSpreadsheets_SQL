from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import date
import smtplib
import config
import os

def email():
    today = date.today()

    for email in config.email_send:
        msg = MIMEMultipart()
        msg['From'] = config.email_user
        msg['To'] = email
        msg['Subject'] = 'Missing Deficiencies List - ' + str(today)


        body = 'Please see attachment.\n \nNOTE: Regardless of if the script breaks, the most recent deficiencienes file will be attached. So be sure to verify in the log file that there were no CRITICAL errors.\n \nIf there are errors, please send a message to Catherine Egboh or Sergei Serlin'
        msg.attach(MIMEText(body,'plain'))


        today_file = "logfile-{0}.txt".format(str(today))
        log = os.path.join('log_file', today_file)
        deficiency = "missing_deficiencies.xlsx"
        filename_list = [deficiency, log]

        for filename in filename_list:
            attachment = open(filename,'rb')
            part = MIMEBase('application','octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',"attachment; filename= "+filename)

            msg.attach(part)
            text = msg.as_string()
            server = smtplib.SMTP('smtp.office365.com',587)
            server.starttls()
            server.login(config.email_user,config.email_password)


        server.sendmail(config.email_user,config.email_send,text)
        server.quit()