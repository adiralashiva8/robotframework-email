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
            self.COMPANY_NAME = BuiltIn().get_variable_value("${COMPANY_NAME}")

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
         self.date_now, self.total_time, self.COMPANY_NAME)

def send_email(send_email, subject, smtp, from_user, pwd, to, cc, total, passed, failed, percentage, exe_date, elapsed_time, company_name):

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
                .rf-box {
                    max-width: 60%%;
                    margin: auto;
                    padding: 30px;
                    border: 3px solid #eee;
                    box-shadow: 0 0 10px rgba(0, 0, 0, .15);
                    font-size: 16px;
                    line-height: 28px;
                    font-family: 'Helvetica Neue', 'Helvetica', Helvetica, Arial, sans-serif;
                    color: #555;
                }
                
                .rf-box table {
                    width: 100%%;
                    line-height: inherit;
                    text-align: left;
                }
                
                .rf-box table td {
                    padding: 5px;
                    vertical-align: top;
                    width: 50%%;
                    text-align: center;
                }
                
                .rf-box table tr.heading td {
                    background: #eee;
                    border-bottom: 1px solid #ddd;
                    font-weight: bold;
                    text-align: left;
                }
                
                .rf-box table tr.item td {
                    border-bottom: 1px solid #eee;
                }
            </style>
        </head>
        <body>

            <div class="rf-box">
                <table cellpadding="0" cellspacing="0">
                    <tr class="top">
                        <td colspan="2">
                            <table>
                                <tr>
                                    <td></td>
                                    <td style="text-align:middle">
										<h1>%s</h1>
									</td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>

                <p style="padding-left:20px">
                    Hi Team,<br>
                    Following are the last build execution result. Please refer <a href="">report.html</a> for more info
                </p>

                <table style="width:80%%;padding-left:20px">
                    <tr class="heading">
                        <td>Test Status:</td>
                        <td></td>
                    </tr>
                    <tr class="item">
                        <td>Total</td>
                        <td>%s</td>
                    </tr>
                    <tr class="item">
                        <td>Pass</td>
                        <td style="color:green">%s</td>
                    </tr>
                    <tr class="item">
                        <td>Fail</td>
                        <td style="color:red">%s</td>
                    </tr>
                </table>

                <br>

                <table style="width:80%%;padding-left:20px">
                    <tr class="heading">
                        <td>Other Info:</td>
                        <td></td>
                    </tr>
                    <tr class="item">
                        <td>Pass Percentage (%%)</td>
                        <td>%s</td>
                    </tr>
                    <tr class="item">
                        <td>Executed Date</td>
                        <td>%s</td>
                    </tr>
                    <tr class="item">
                        <td>Machine</td>
                        <td>%s</td>
                    </tr>
                    <tr class="item">
                        <td>OS</td>
                        <td>%s</td>
                    </tr>
                    <tr class="item">
                        <td>Duration</td>
                        <td>%s</td>
                    </tr>
                </table>

                <table>
                    <tr>
                        <td style="text-align:center;color: #999999; font-size: 11px">
                            <p>Best viewed in websites!</p>
                        </td>
                    </tr>
                </table>
            </div>
        </body>

        </html>
        """ % (company_name, total, passed, failed, percentage, exe_date, platform.uname()[1], platform.uname()[0], elapsed_time)

        msg.attach(MIMEText(email_content, 'html'))

        server.starttls()
        server.login(msg['From'], pwd)
        server.sendmail(from_user, to_addrs, msg.as_string())