import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import re
import pandas as pd
import numpy as np
from numpy import nan

mail_host = 'smtp.126.com'
mail_user = 'abcdeg2023@126.com'
mail_pass = 'TAJQMAJQEIWDZRLO'
sender = 'abcdeg2023@126.com'

t1=pd.read_excel("各队电子邮箱1.xlsx",sheet_name=0)
for i in range(433,465):
    receivers = []
    for j in range(3,8):
        value=t1.iloc[i, j]
        receivers.append(value)#增加收件人

    t=re.sub(r'\D', '', str(t1.iloc[i, 0]))
    tmp=t+'.pdf'#选择发送的pdf文件

    # receivers.remove('nan')

    receivers = list(filter(lambda x: x is not nan and x is not None, receivers))
    # newlist = [x for x in mylist if pd.isnull(x) == False]
    # print(receivers)
    # print(len(receivers))

    if(os.path.exists(tmp)):
        # 设置eamil信息
        message = MIMEMultipart()
        message['From'] = sender
        message['To'] = ','.join(receivers)
        message['Subject'] = '4C奖状'
        part = MIMEApplication(open(tmp,'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename=tmp)
        message.attach(part)
        # 登录并发送
        try:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(mail_host, 25)
            smtpObj.login(mail_user, mail_pass)
            smtpObj.sendmail(
                sender, receivers, message.as_string())
            print('success' + str(receivers) + ' ' + tmp)
            smtpObj.quit()
        except smtplib.SMTPException as e:
            print("wrong" + str(receivers) + ' ' + tmp)
            print('error', e)
            for i in range(0, len(receivers)):
                try:
                    smtpObj = smtplib.SMTP()
                    smtpObj.connect(mail_host, 25)
                    smtpObj.login(mail_user, mail_pass)
                    smtpObj.sendmail(
                        sender, receivers[i], message.as_string())
                    print('success' + str(receivers[i]) + ' ' + tmp)
                    smtpObj.quit()
                except smtplib.SMTPException as e:
                    print(receivers[i]+"二次重发失败")
    else:
        print("不存在"+tmp)
