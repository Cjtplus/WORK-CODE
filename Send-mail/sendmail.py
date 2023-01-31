from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def sendmail(attachment=None):
    # 构造邮件
    # 格式化
    msg = MIMEMultipart()

    # 邮件头部
    # 邮件标题
    mail_title = '发送邮件'
    msg['Subject'] = Header(mail_title, 'utf-8')
    # 登录账号密码 以139邮箱为例
    sender = '**@139.com'
    password = ''  # 邮箱密码/执行码
    # 发件人
    msg['From'] = '' + '<' + sender + '>'
    # 收件人
    receiver = ''
    # receiver = [] # 有多个收件人时执行
    msg['To'] = '' + '<' + receiver + '>'  # 单个收件人时执行
    # msg['To'] = ";".join('' + '<' + receiver + '>') # 多个收件人时执行

    # 正文内容
    content = """发送邮件"""

    """
    可以增加HTML语句对邮件正文内容进行各种格式操作，不涉及css
    
    <div style="font-family: 微软雅黑;font-size: 16.0px;color: rgb(0,0,0)">发送邮件</div>
    
    但可能在部分浏览器中失效
    """
    """
    邮件内容涉及表格：
    原数据为Dataframe格式
    content_table_data = ''
    for row in range(1, len(df) + 1):
        td = ''
        for col in range(len(df.columns)):
            cellData = df[row - 1][col]
            if row == len(df):
                tip = '<td>' + str(cellData) + '</td>'
                td = td + tip
                tr = '<tr>' + td + '</tr>'
            else:
                # 读取单元格数据，赋给cellData变量供写入HTML表格中
                tip = '<td>' + str(cellData) + '</td>'
                td = td + tip
                tr = '<tr>' + td + '</tr>'
            # tr = tr.encode('utf-8')
        content_table_data = content_table_data + tr
        
    同样的，可以增加HTML语句对表格内容进行美化
    """

    # content_table = content + content_table_data
    msg.attach(MIMEText(content, _subtype='html', _charset='utf-8'))  # 邮件正文内容

    if attachment is None:
        pass
    else:
        # 添加附件
        xlsx_file = ''  # 附件路径
        xlsxApart = MIMEApplication(open(xlsx_file, 'rb').read())
        xlsxApart.add_header('Content-Disposition', 'attachment', filename='')

        msg.attach(xlsxApart)

    # 发送
    host_server = 'smtp.139.com'  # 139邮箱的smtp地址
    smtp = SMTP_SSL(host_server)  # SSL 登录
    smtp.set_debuglevel(0)
    smtp.ehlo(host_server)
    smtp.login(sender, password)
    smtp.sendmail(sender, receiver, msg.as_string())
    smtp.quit()
    print("邮件发送成功")
