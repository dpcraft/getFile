
#coding:utf-8

import poplib
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import os
from Tkinter import Tk
from tkMessageBox import showwarning
import win32com.client as win32
import sys
reload(sys)
sys.setdefaultencoding('utf8')

def mkdir(path):
    path = path.strip()
    path = path.rstrip('\\')
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        return True
    else:
        print unicode(path,"gb2312") + '目录已存在'
        return False

def decode_str(s):
    if not s:
        return None
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def get_mails(prefix):
    host = 'pop.163.com'

    username = 'username'
    #username不包含@163.com
    password = 'password'

    server = poplib.POP3(host)
    try:
        server.user(username)
        server.pass_(password)
    except poplib.error_proto,e:
        print "Login filed:" + e.message
        sys.exit(1)

    warn = lambda app: showwarning(app,"完成?")
    app = 'Excel'
    x1 = win32.gencache.EnsureDispatch('%s.Application' % app)
    ss = x1.Workbooks.Add()
    sh = ss.ActiveSheet
    x1.Visible = False
    path = "f:\\"+ "数据库实验报告".encode("gb2312") + "\\" + folder.encode("gb2312") + "\\"
    mkdir(path)

    # 获得邮件
    messages = [server.retr(i) for i in range(1, len(server.list()[1]) + 1)]
    messages = [b'\r\n'.join(mssg[1]).decode() for mssg in messages]
    messages = [Parser().parsestr(mssg) for mssg in messages]
    print("===="*10)
    messages = messages[::-1]
    rownum = 1
    mailNO = 0
    for message in messages:
        subject = message.get('Subject')
        subject = decode_str(subject)
        mailNO = mailNO + 1
        #如果标题匹配
        if subject: #and subject[-3:] == prefix:
            value = message.get('From')
            if value:
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = u'%s <%s>' % (name, addr)
            # print("发件人: %s" % value)
            # print("标题:%s" % subject)
            for part in message.walk():
                fileName = part.get_filename()
                fileName = decode_str(fileName)
                if fileName:
                    stdInfo = fileName.split('_')
                # 保存附件
                #     print stdInfo[0][0:4]
                #     print stdInfo[0][0:4]==str(2015)
                if fileName and (stdInfo.__len__() == 3) and  (stdInfo[0][0:4] == str(2015)) and (stdInfo[2].split('.')[0][-3:] == prefix):
                    if os.path.exists(unicode(path,"gb2312") + fileName):
                        print fileName + "已存在"
                    else:
                        with open(unicode(path,"gb2312") + fileName, 'wb') as fEx:
                            data = part.get_payload(decode=True)
                            fEx.write(data)
                            print "附件%s已保存" % fileName
                    print ("----" * 10)
                    sh.Cells(rownum,1).Value = "'" + stdInfo[0]
                    sh.Cells(rownum,2).Value = stdInfo[1]
                    sh.Cells(rownum,3).Value = stdInfo[2].split('.')[0].encode('gb2312')
                    rownum = rownum + 1
    print "邮件总数为: " + str(mailNO)
    server.quit()
    warn(app)
    if os.path.exists(unicode(path,"gb2312") + folder + "统计" +".xlsx"):
        os.remove(unicode(path,"gb2312") + folder + "统计" +".xlsx")

    ss.SaveAs(unicode(path,"gb2312") + folder + "统计" +".xlsx")
    ss.Close()
    x1.Application.Quit()

if __name__ == '__main__':
    Tk().withdraw()
    prefix = str("实验一")
    folder = str("实验一tao")
    get_mails(prefix)
