# *===================================*
# -*- coding: utf-8 -*-
# * Time : 2019/11/29 10:39
# * Author : zhangsf
# *===================================*
from importlib import reload

import win32com.client as win32
import warnings
import pythoncom
import sys
reload(sys)
warnings.filterwarnings('ignore')
pythoncom.CoInitialize()
def sendmail():
    #邮件主题
    sub = 'outlook python mail test'
    #邮件的body
    body = 'my test\r\n my python mail'
    outlook = win32.Dispatch('outlook.application')
    #修改下面的接收邮箱
    receivers = ['XXX@XX.com']
    mail = outlook.CreateItem(0)
    mail.To = receivers[0]
    mail.Subject = sub.encode('utf-8').decode('utf-8')
    mail.Body = body.encode('utf-8').decode('utf-8')
    #修改需要发送的邮件的附件信息
    mail.Attachments.Add('E:\code\python\SendEmail\/text1.txt')
    mail.Send()
sendmail()