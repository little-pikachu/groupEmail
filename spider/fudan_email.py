import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import xlrd

class group_email():

    def __init__(self):
        chrome_oprions = Options()
        chrome_oprions.add_argument("--headless")  # 使用无头浏览器，即不显示网页
        self.driver = webdriver.Chrome(options=chrome_oprions,executable_path='./chromedriver_win32/chromedriver.exe')

    def login(self, name, password, url='https://mail.fudan.edu.cn/'):
        print('登录中')
        time.sleep(0.5)
        self.driver.get(url=url)
        el_username = self.driver.find_element_by_id('uid')
        el_password = self.driver.find_element_by_id('password')
        el_username.send_keys(name)
        el_password.send_keys(password)
        el_login_btn = self.driver.find_element_by_xpath("//div[@class='inpB']/button[@class='Button']")
        el_login_btn.click()
        time.sleep(2)  # wait for two seconds

    def read_excel(self,file,sheet_index=0):
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(sheet_index)
        print('读取表格文件')
        # print("工作表名称:", sheet.name)
        # print("行数:", sheet.nrows)
        # print("列数:", sheet.ncols)
        data = []
        headers = sheet.row_values(0)
        for i in range(1, sheet.nrows):
            email = {}
            for index in range(len(headers)):
                value = sheet.row_values(i)[index]
                if type(value) == type(0.0):
                    value = int(value)
                email[headers[index]] = str(value)
            data.append(email)
        return data

    def send_email2(self, path, title, message_template, server=''):  # 发送多份邮件，不同的人邮件信息不一样
        emails = self.read_excel(path)
        print(emails)
        if len(emails) == 0:
            return
        if 'email' not in  emails[0]:
            print('excel文件格式不正确，不存在列excel')
            return
        if emails[0]['email'].find('@') == -1 and server == '':
            print('数据表格中邮箱地址不存在服务器信息，请传入send_email1的第四个参数server，格式如server="@fudan.edu.cn"')
            return
        for i in range(len(emails)):
            email = emails[i]
            self.send_email([email], title, message_template.format(config=email), i+1, server)

    def send_email1(self, path,  title, message, server=''):   # 给所有人发送一样的信息只需要在接收人处填写多个地址就可
        emails = self.read_excel(path)
        print(emails)
        if len(emails) == 0:
            return
        if 'email' not in  emails[0]:
            print('excel文件格式不正确，不存在列excel')
            return
        if emails[0]['email'].find('@') == -1 and server == '':
            print('数据表格中邮箱地址不存在服务器信息，请传入send_email2的第四个参数server，格式如server="@fudan.edu.cn"')
            return

        self.send_email(emails, title, message, 1, server)

    def send_email(self,emails,title,message,compose,server):
        print('发送邮件给以下用户')
        el_write = self.driver.find_element_by_xpath("//a[@totabid='compose' and @class='compose']")
        el_write.click()
        iframe = 'compose' + str(compose)
        self.driver.switch_to.frame(iframe)
        time.sleep(1)
        el_input = self.driver.find_element_by_css_selector("#InputEx_To").find_element_by_css_selector("input.inputElem")

        #输入收件人地址
        if type(emails) == type([]):
            for item in emails:
                mail = item['email'] + server + ';'
                el_input.send_keys(mail)
                print(mail, end=' ')
                time.sleep(0.27)
        else:
            return
        #输入主题
        el_title = self.driver.find_element_by_id('subject')
        el_title.send_keys(title)

        #输入内容
        self.driver.switch_to.frame('htmleditor')
        self.driver.switch_to.frame('HtmlEditor')
        js = 'let span = document.querySelector("#spnEditorSign") ===null?"":document.querySelector("#spnEditorSign").outerHTML;document.querySelector("body").innerHTML = "' + message + '" + "<br/><br/><br/>" + span;'
        self.driver.execute_script(js)
        #发送按钮
        self.driver.switch_to.default_content() # 切换回主文档
        self.driver.switch_to.frame(iframe)
        btn_send = self.driver.find_element_by_id('btnSend')
        btn_send.click()
        print('发送完成')
        self.driver.switch_to.default_content()
        time.sleep(0.5)


email = group_email()
email.login(name='17302010049',password='password')
path = '../data/emails.xlsx'

# 第一种群发方式为：对于所有的用户发送一样的数据内容。
title='测试--群发'
message='这是一段测试数据'
email.send_email1(path, title, message,server="@fudan.edu.cn")

# 第二种群发方式为：对于不同的用户发送不同的数据内容，但是title是一样。message_template为数据模板
# title = '华为云账号通知--群发'
# message_template = '同学你好，你的华为云账号为{config[hwy_account]}，密码为{config[hwy_password]}' # excel中的列名为hwy_account和hwy_password
# email.send_email2(path,title,message_template,server="@fudan.edu.cn")
