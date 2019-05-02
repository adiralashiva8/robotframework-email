import os
import smtplib
import math
import datetime
import platform
import time

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from robot.libraries.BuiltIn import BuiltIn

class EmailListener:

    ROBOT_LISTENER_API_VERSION = 2

    def __init__(self):
        self.total_tests = 0
        self.passed_tests = 0
        self.failed_tests = 0
    
        self.PRE_RUNNER = 0
        self.start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
    
    def start_suite(self, name, attrs):

        # Fetch email details only once
        if self.PRE_RUNNER == 0:

            self.SEND_EMAIL = BuiltIn().get_variable_value("${SEND_EMAIL}")
            self.SMPT = BuiltIn().get_variable_value("${SMPT}")
            self.SUBJECT = BuiltIn().get_variable_value("${SUBJECT}")
            self.FROM = BuiltIn().get_variable_value("${FROM}")
            self.PASSWORD = BuiltIn().get_variable_value("${PASSWORD}")
            self.TO = BuiltIn().get_variable_value("${TO}")
            self.CC = BuiltIn().get_variable_value("${CC}")

            self.PRE_RUNNER = 1

            self.date_now = datetime.datetime.now().strftime("%Y-%m-%d")

        self.test_count = len(attrs['tests'])
    
    def end_test(self, name, attrs):
        if self.test_count != 0:
            self.total_tests += 1
        
        if attrs['status'] == 'PASS':
            self.passed_tests += 1
        else:
            self.failed_tests += 1
    
    def close(self):
        self.end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
        self.total_time=(datetime.datetime.strptime(self.end_time,'%H:%M:%S') - datetime.datetime.strptime(self.start_time,'%H:%M:%S'))

        send_email(self.SEND_EMAIL, self.SUBJECT, self.SMPT, self.FROM, self.PASSWORD, self.TO, self.CC,
         self.total_tests, self.passed_tests, self.failed_tests, math.ceil(self.passed_tests * 100.0 / self.total_tests),
         self.date_now, self.total_time)

def send_email(send_email, subject, smtp, from_user, pwd, to, cc, total, passed, failed, percentage, exe_date, elapsed_time):

    if send_email:
        server = smtplib.SMTP(smtp)
    
        msg = MIMEMultipart()
        msg['Subject'] = subject

        msg['From'] = from_user
        msg['To'] = to
        msg['Cc'] = cc
        to_addrs = [to] + [cc]
        msg.add_header('Content-Type', 'text/html')

        email_content = """
        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
        <title>Automation Status</title>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=edge" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0 " />
            <style>
                body {
                    background-color:#F2F2F2; 
                }
                body, html, table,span,b {
                    font-family: Courier New;
                    font-size: 14px;
                }
                .pastdue { color: crimson; }
                table {
                    border: 1px solid silver;
                    padding: 6px;
                    margin-left: 30px;
                    width: 300px;
                }
                tbody {
                    text-align: center;
                }
                td {
                    height: 25px;
                }
                .dt-buttons {
                    margin-left: 30px;
                }
            </style>
        </head>
        <body>
        <span>Hi Team,<br><br>Following are the last build execution status.<br><br>Result:<br><br></span>
            <table>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">Total</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">Pass</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">Fail</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">Pass Percentage</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">Execution Date</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">Machine</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">OS</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                <tr>
                   <td style="background-color: #DCDCDC;text-align: center;">Duration</td>
                   <td style="text-align: center;">%s</td> 
                </tr>
                </tbody>
            </table>

            
            <span><br>Regards,<br>QA Team</span>

        </body></html> 
        """ % (total, passed, failed, percentage, exe_date, platform.uname()[1], platform.uname()[0], elapsed_time)

        msg.attach(MIMEText(email_content, 'html'))

        server.starttls()
        server.login(msg['From'], pwd)
        server.sendmail(from_user, to_addrs, msg.as_string())