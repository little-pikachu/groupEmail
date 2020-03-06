import smtplib
from email.mime.text import MIMEText
from email.header import Header

# 发信服务器
#smtp_server = 'mail.fudan.edu.cn' 授权码就是密码
#smtp_server = 'smtp.qq.com' 授权码需要去qq邮箱设置

from_addr = input('请输入登录邮箱：')
smtp_server = input('发信服务器：')
password = input('请输入邮箱授权码：')


# 邮件内容
text = '''
hey 这是我用Python发的第一封邮件
人生苦短，我用Python
'''

emails = [{'email':'1260968291@qq.com'}, {'email':'17302010049@fudan.edu.cn'}]

server = smtplib.SMTP_SSL(host=smtp_server)
server.connect(smtp_server, port=465)
server.login(from_addr, password)

for email in emails:
    msg = MIMEText(text, 'plain', 'utf-8')
    msg['From'] = Header(from_addr)
    msg['To'] = Header(email['email'])
    msg['Subject'] = Header('python test')
    server.sendmail(from_addr, email['email'], msg.as_string())

# 关闭服务器
server.quit()
