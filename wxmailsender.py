#!/usr/bin/python
#-*- coding:utf-8 -*-

import wx

import sys

class wxmailsender(wx.Frame):

    def __init__(self, title):
        wx.Frame.__init__(self, None, -1, title=title,
                                           size=(400, 230)) 
        self.excelfile = ''
        self.InitUI()
        self.Centre()
        self.Show()

    def InitUI(self):
        self.panel = wx.Panel(self, -1)
        self.fileNameInput = wx.TextCtrl(self.panel, -1, u"", pos=(35, 40),size=(260,25))
        self.selectFile = wx.Button(self.panel, label=u"选择", pos=(300 ,40), size=(60, 25))
        self.selectFile.Bind(wx.EVT_BUTTON, self.SelectFile)
        self.fafangRadioButton = wx.RadioButton(self.panel, label=u"发放说明", pos=(35,90), style=wx.RB_GROUP)
        self.jisongRadioButton = wx.RadioButton(self.panel, label=u"寄送通知", pos=(150, 90))
        self.fafangRadioButton.Bind(wx.EVT_RADIOBUTTON, self.SetMailType)
        self.jisongRadioButton.Bind(wx.EVT_RADIOBUTTON, self.SetMailType)
        
        self.sendButton = wx.Button(self.panel, label=u'发送', pos=(300, 120), size=(60,30))
        self.sendButton.Bind(wx.EVT_BUTTON, self.SendMail)
        
        self.showState = wx.StaticText(self.panel, -1, u"", pos=(90, 155), size=(120, 20))
        self.showState.SetForegroundColour('blue')

    def SendMail(self, event):
        import parseExcelAndEmailSender
        
        if self.fafangRadioButton.GetValue():
            mailTemplate = parseExcelAndEmailSender.fafangshuomingFile
            dataItems = parseExcelAndEmailSender.fafangshuomingDataItems
            mail_subject = "一起公益网每月公益券发放说明"
        elif self.jisongRadioButton.GetValue():
            mailTemplate = parseExcelAndEmailSender.jisongtongzhiFile
            dataItems = parseExcelAndEmailSender.jisongtongzhiDataItems
            mail_subject = "一起公益网每月公益券寄送通知"

        if self.excelfile == '':
            self.showState.SetLabel(u"未选择Excel文件")
        else:
            pas = parseExcelAndEmailSender.parseAndSend(self.excelfile, mailTemplate, dataItems, mail_subject)
            pas.processExcel()
            pas.sendResultToAdminMail()

    def SelectFile(self, event):
        selectFileDialog = wx.FileDialog(self, message=u"请选择Excel文件",
                                        wildcard=u"*xls",
                                        style=wx.SAVE)
        self.excelfile = u''
        if selectFileDialog.ShowModal() == wx.ID_OK:
            self.excelfile = selectFileDialog.GetPath()
        self.fileNameInput.SetValue(self.excelfile)
    
    def SetMailType(self, event):
        if self.fafangRadioButton.GetValue():
            self.showState.SetLabel(u'您选择了"发放说明"')
        if self.jisongRadioButton.GetValue():
            self.showState.SetLabel(u'您选择了"寄送通知"')

if __name__ == '__main__':
    app = wx.App()
    wxmailsender(title=u"一起公益网---邮件自动发送器")
    app.MainLoop()
