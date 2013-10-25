## 依赖

1.
wxPython

2.
xlrd

## 关于配置文件config.py

config.py文件中有三个选项：debug, debugMail, adminMail。相关说明请参考config.py内的注释。

## Excel格式
使用之前请确保你要发送的Excel符合预定义的格式。

1). "发放说明"的Excel格式如下:

姓名	月	开始日	月	截止日	有效提交抵消后条数	公益活动信息	公益组织信息	张数	含额外奖励张数	点数（张数*50点/张）	邮箱


2). "寄送通知"的Excel格式如下:

姓名	月	开始日	月	截止日	有效提交抵消后条数	公益活动信息	公益组织信息	张数	含额外奖励张数	点数（张数*50点/张）	寄送月	寄送日	寄送方式	寄送单号	邮箱

### 使用

1.
命令行：`python parseExcelAndEmailSender.py -f filename` 或 `python parseExcelAndEmailSender.py -j filename`。

2.
图形界面： `python wxmailsender.py`
