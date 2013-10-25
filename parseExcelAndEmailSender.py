#!/usr/bin/env python
#-*- coding: utf-8 -*-

import xlrd
import sys
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import time
import config

reload(sys)
sys.setdefaultencoding('utf-8')

fafangshuomingFile = "fafangshuoming.txt"
jisongtongzhiFile = "jisongtongzhi.txt"

fafangshuomingDataItems = ['姓名', '开始月', '开始日', '截止月', '截止日', '有效提交抵消后条数', '公益活动信息', '公益组织信息', '张数', '含额外奖励张数', '点数', '邮箱']
jisongtongzhiDataItems = fafangshuomingDataItems[0:-1]
jisongtongzhiDataItems.extend(['寄送月', '寄送日' ,'寄送方式', '寄送单号'])
jisongtongzhiDataItems.append(fafangshuomingDataItems[-1])

class parseAndSend():
    def __init__(self,  excelfile, mailTemplate, dataItems, mail_subject):
        self.logFile = "parseExcel.log"
        self.excelFile = excelfile
        self.mail_host = "smtp.gmail.com:587"
        self.mail_user = "xxx@gmail.com"
        self.mail_pass = "xxx"
        self.mail_user_toshow = "系统自动邮件，请勿回复" + "<" + self.mail_user + ">"
        self.adminMail = config.adminMail
        self.sendResults = list()

        self.mailTemplate = mailTemplate
        self.dataItems = dataItems
        self.mail_subject = mail_subject

    def writeLog(self, message):
        with open(self.logFile, 'a') as logFileHandler:
            logFileHandler.write(message)

    def sendMail(self, content, mailto, subject):
        if config.debug:
            mailto = config.debugMail
        msg = MIMEText(content, 'html', 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = Header(self.mail_user_toshow, 'utf-8')
        msg['To'] = mailto
        if config.debug:
            print "True Mail is %s" % mailto
        try:
            s = smtplib.SMTP(self.mail_host)
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(self.mail_user, self.mail_pass)
            s.sendmail(self.mail_user_toshow, mailto, msg.as_string())
            s.quit()
            return True
        except Exception, e:
            print str(e)
            return False
    
    def sendRowInfoMail(self, mailTemplateContent, data):
        contentData = list()
        for item in self.dataItems[:-1]:
            contentData.append(data[item])
        contentData = tuple(contentData)
        content = mailTemplateContent % contentData
        mailto = data['邮箱']
        print "sending to %s" % mailto
        if self.sendMail(content, mailto, self.mail_subject):
            successLog = "send to %s(%s) success!" % (data['姓名'], mailto)
            self.writeLog(successLog+"\n")
            self.sendResults.append(successLog)
        else:
            failLog = "send to %s(%s) failed!" % (data['姓名'], mailto)
            self.writeLog(failLog+"\n")
            self.sendResults.append(failLog)

    def processExcel(self):
        self.writeLog("\n----------"+time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())+"------------\n")
        bk = xlrd.open_workbook(self.excelFile)
        sh = bk.sheet_by_index(0)
        nrows = sh.nrows

        row_list = []
        for i in range(1, nrows):
            row_data = sh.row_values(i)
            row_list.append(row_data)
        for rowvalue in row_list:
            innerindex = 0
            dataDict = dict()
            for element in rowvalue:
                if isinstance(element, float):
                    element = int(element)
                element = str(element)
                self.writeLog(element + "   ")
                dataDict[self.dataItems[innerindex]] = element
                innerindex += 1
            self.writeLog("\n")
            with open(self.mailTemplate) as fileHandler:
                content = fileHandler.read()
                self.sendRowInfoMail(content, dataDict)

    def sendResultToAdminMail(self):
        resultSubject = "%s %s情况汇总" % (time.strftime("%Y-%m-%d", time.localtime()), self.mail_subject)
        content = ''
        print len(self.sendResults)
        for item in self.sendResults:
            content += item
            content += "<br />"

        self.sendMail(content, self.adminMail, resultSubject)

if __name__ == '__main__':
    
    excelfile = sys.argv[2]

    if sys.argv[1] == '-f':
        mailTemplate = fafangshuomingFile
        dataItems = fafangshuomingDataItems
        mail_subject = "一起公益网每月公益券发放说明"
    elif sys.argv[1] == '-j':
        mailTemplate = jisongtongzhiFile
        dataItems = jisongtongzhiDataItems
        mail_subject = "一起公益网每月公益券寄送通知"
    
    pas = parseAndSend(excelfile, mailTemplate, dataItems, mail_subject)
    pas.processExcel()
    for index in xrange(0, 5):
        if pas.sendResultToAdminMail():
            break
