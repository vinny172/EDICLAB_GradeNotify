#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""
專案：EDIC_LAB成績自動發送系統
用途：透過EXCEL匯入學生EMAIL和成績資料，本系統可將個人成績發送至各別的Email信箱。
備註：
版本：1.3
開發環境：Python 2.7
開發人：Colin Lin
"""

import base64, textwrap, traceback, smtplib, mimetypes, xlrd, ttk
from Tkinter import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

version = "1.3" #版本

class GUI_Window(Frame):
    
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()
        self.createWidgets()
        self.loadconfig()     
        self.loadexcel()
        
    def loadconfig(self):
        try:
            config_file = open('cfg.ini','r')
        except:
            return
        self.loadlist = config_file.read().splitlines()      
        while len(self.loadlist) < 5:
            self.loadlist.append("")
        self.loadlist[1] = base64.decodestring(self.loadlist[1])
        config_file.close()
        self.default_account.set(self.loadlist[0])
        self.default_password.set(self.loadlist[1])
        self.SMTPList.current(int(self.loadlist[2]))
        self.default_testmail.set(self.loadlist[3])
        
    def loadexcel(self):
        self.data = xlrd.open_workbook('./Grades.xlsx')
        self.student = self.data.sheet_by_name(u'學生通訊錄')
        self.score_sheet = self.data.sheet_by_name(u'成績表')
        self.setup = self.data.sheet_by_name(u'設定')

        self.year = int(self.setup.cell_value(1,2))
        self.semester = int(self.setup.cell_value(2,2))
        self.course = self.setup.cell_value(3,2)
        self.exam = self.setup.cell_value(4,2)
        self.default_year.set(self.year)             #從EXCEL讀出之年度
        self.default_semester.set(self.semester)      #從EXCEL讀出之學期
        self.default_course.set(self.course)        #從EXCEL讀出之課程
        self.default_exam.set(self.exam)          #從EXCEL讀出之考試名稱
        self.StudentCountText["text"] = u"學生人數：" + str(self.student.nrows-2)
        self.StatusCountText["text"] = "載入完成！"
        


    def sendmail(self, to, subject, textbody):
        
        username = self.accountField.get()
        password = self.passwordField.get()
        server, port = ['',0]
        
        if self.SMTPList.current() == 0:
            server = 'smtpauth.net.nthu.edu.tw'    #清大SMTP伺服器
            port = 25
        else:
            server = 'smtp.gmail.com'         #Gmail SMTP伺服器
            port = 587  

        try:
            server = smtplib.SMTP(server, port)
            self.StatusCountText["text"] = "SMTP伺服器連線成功"
            server.ehlo()
            server.starttls()
            server.login(username,password)  # 若SMTP server 不需要登入則可省略
            self.StatusCountText["text"] = "成功登入"
            msg = MIMEMultipart() 
            msg["From"] = username
            msg["To"] = to 
            msg["Subject"] =  subject
            msg["preamble"] = u'不支援檢視MIME信件.\n' 
            part = MIMEText(textbody, _charset="UTF-8") 
            msg.attach(part)
            server.sendmail(username, to, msg.as_string())  #發送信件
            server.quit()
        except:
            tp, val, tb = sys.exc_info()
            self.StatusCountText["text"] = u"寄送失敗 錯誤：" + str(val)
        else:        
            self.StatusCountText["text"] = u"成功寄出「" + subject + u"」"

    def clickhelp(self):
        helpwin = Toplevel(root)
        helpwin.title(u"說明")
        help_text = Text(helpwin, height = 12,width = 40)
        help_text.insert(INSERT, textwrap.dedent('''\
            說明：
            1)於EXCEL檔中，填入設定分頁相關資訊
            2)將學生EMAIL填入通訊錄分頁
            3)將分數填入分數表分頁
            4)於本程式填入SMTP連線資訊並測試寄信
            5)於測試信箱確認是否收到信
            6)確認相關資訊無誤即可開始發送

            ※使用清大信箱帳號，帳號需填完整EMAIL
            (ex.帳：s1040635xx@m104.nthu.edu.tw)
            ※注意：若使用GOOGLE帳號，需設置帳戶安全性相關設定，才可使用GMAIL登入
            
        '''))
        help_text.pack()
        button = Button(helpwin, text=u"關閉" ,command=helpwin.destroy)
        button.pack()

    def clickabout(self):
        aboutwin = Toplevel(root)
        aboutwin.title(u"關於")
        title = Label(aboutwin, text=u"EDIC_LAB成績通知系統V" + version)
        title.config(font=("Microsoft JhengHei",16))
        body = Label(aboutwin, text=u"版本："+ version +u"\n開發環境：Python 2.7\n開發者：Colin Lin" , justify="left")
        date = Label(aboutwin, text = "2017/05/08")
        button = Button(aboutwin, text=u"關閉" ,command=aboutwin.destroy)
        title.grid(row=0,column=0)
        body.grid(row=1,column=0,sticky=W)
        date.grid(row=2,column=0,sticky=E)
        button.grid(row=3,column=0)

    def clicktest(self):
        testmail = self.testmailField.get()
        title = u'EDIC_LAB成績通知系統_測試信件'
        body = u"測試成功！！！！"
        self.sendmail(testmail, title, body)
        
    def list_catch(self, index):
        ID = unicode(self.student.cell_value(index+2,1)).split('.')[0]
        name = unicode(self.student.cell_value(index+2,2))
        mail = unicode(self.student.cell_value(index+2,3))
        score = unicode(self.score_sheet.cell_value(index+2,3)).split('.')[0]
        return [ID,name,mail,score]

    def item_status(self, status):
        selected_item = self.preview_list.selection()
        origin_data = list(self.list_catch(int(selected_item[0])-1))
        origin_data.append(status)
        self.preview_list.item(selected_item, values=tuple(origin_data))
        return

    def del_items(self):
        for i in self.preview_list.selection():
            self.preview_list.delete(i)
        return

    def preview(self):
        self.previewwin = Toplevel(root)
        self.previewwin.title(u"預覽視窗("+ (self.course) +  (self.exam) + u")" )
        self.preview_list = ttk.Treeview(self.previewwin, show="headings", height=18, columns=("ID", "NAME", "EMAIL", "SCORE", "STATUS"))
        for i in xrange(self.student.nrows-2):
            self.preview_list.insert("", "end", values=tuple(self.list_catch(i)), iid=(i+1))
        
        self.preview_list.column("ID", width=100, anchor="center")
        self.preview_list.column("NAME", width=90, anchor="center")
        self.preview_list.column("EMAIL", width=280, anchor="center")
        self.preview_list.column("SCORE", width=50, anchor="center")
        self.preview_list.column("STATUS", width=50, anchor="center")
        self.preview_list.heading("ID", text="學號")
        self.preview_list.heading("NAME", text="姓名")
        self.preview_list.heading("EMAIL", text="EMAIL")
        self.preview_list.heading("SCORE", text="分數")
        self.preview_list.heading("STATUS", text="狀態")
       
        vbar = ttk.Scrollbar(self.previewwin, orient=VERTICAL, command=self.preview_list.yview)
        self.preview_list.configure(yscrollcommand=vbar.set)
        del_button = Button(self.previewwin, text=u"刪除項目" , command = self.del_items)
        test_button = Button(self.previewwin, text=u"test" , command = self.item_status)
        send_button = Button(self.previewwin, text=u"寄送選取部分" , command = self.clicksend)
        
        start_send_all = Button(self.previewwin, text=u"寄送全部" , command = self.clicksend_all)
        close_button = Button(self.previewwin, text=u"取消" , command = self.previewwin.destroy)
        
        self.preview_list.grid(row=0, column=0, columnspan=5, sticky=NSEW, padx=0, pady=0)
        vbar.grid(row=0,column=5,columnspan=1, rowspan=2, sticky=S+N)
        start_send_all.grid(row=1, column=0, columnspan=2, sticky=NSEW, padx=0, pady=0)
        close_button.grid(row=1, column=4, columnspan=1, sticky=NSEW, padx=0, pady=0)
        del_button.grid(row=1,column=3,columnspan=1, sticky=NSEW)
        #test_button.grid(row=1,column=1,columnspan=3)
        send_button.grid(row=1,column=2,columnspan=1, sticky=NSEW)

    def clicksend_all(self):
        for i in xrange(self.student.nrows-2):
            try:
                self.preview_list.selection_add(i+1)
            except:
                continue

            if self.preview_list.next(i+1)=="":
                break
        self.clicksend()

    def clicksend(self):
        for i in self.preview_list.selection():
            self.preview_list.selection_set(i)
            ID, name, mail, score = self.preview_list.item(i)["values"][:4]
            title = self.course +  self.exam + u"成績通知_" + str(ID)
            hello = name[0] + u"同學您好：\n\n"
            context = u"　　您於" + str(self.year) + u"學年度第" + str(self.semester) + u"學期所修習的課程" + self.course + u"，" + self.exam + u"已批改完成，您的分數為" + str(score) + u"分，若有疑義麻煩於助教時間至綜二3F EDIC LAB反應，謝謝！\n\nEDICLAB助教群\n\n"
            if ID=="" or name=="" or mail=="" or score=="":
                self.item_status(u"缺")
                continue
            try:
                self.sendmail(mail, title, hello + context)
                #print (mail, title)
                self.item_status(u"完成")
            except:
                self.item_status(u"錯誤")
			
    def saveconfig(self):
        config_file = open('cfg.ini','w')
        if self.checkboxvar.get() == True:
            info_list = [self.accountField.get(),base64.encodestring(self.passwordField.get()).strip(),str(self.SMTPList.current()),self.testmailField.get()]
            config_file.write('\n'.join(info_list))
        elif self.checkboxvar.get() == False:
            config_file.write("")
        config_file.close()

    def click_exit(self):
        self.saveconfig()
        root.destroy()

    def click_test(self):
        print "For Development"

    def createWidgets(self):
        #工具欄
        self.Menubar = Menu(self)
        funcmenu = Menu(self.Menubar, tearoff=0)
        funcmenu.add_command(label=u"離開", command=self.click_exit)
        self.Menubar.add_cascade(label=u"功能", menu=funcmenu)
        Helpmenu = Menu(self.Menubar, tearoff=0)
        Helpmenu.add_command(label=u"使用說明", command=self.clickhelp)
        Helpmenu.add_command(label=u"關於", command=self.clickabout)
        self.Menubar.add_cascade(label=u"說明", menu=Helpmenu)
        root.config(menu=self.Menubar)


        #框架1
        
        self.labelframe1 = LabelFrame(self, text=u"SMTP郵件伺服器設定")
        self.labelframe1.config(width=200)
        self.labelframe1.grid(row=0, column=0, columnspan=4, padx="5",sticky=W)

        self.default_account = StringVar()
        self.default_password = StringVar()
        self.default_testmail = StringVar()

        self.accountText = Label(self.labelframe1)
        self.accountText["text"] = u"帳號："
        self.accountText.grid(row=0, column=0, sticky=E, pady="5")
        self.accountField = Entry(self.labelframe1)
        self.accountField['textvariable'] = self.default_account
        self.accountField["width"] = 20
        self.accountField.grid(row=0, column=1, columnspan=2, sticky=W, pady="5")
 
        self.passwordText = Label(self.labelframe1)
        self.passwordText["text"] = u"密碼："
        self.passwordText.grid(row=1, column=0, sticky=E, pady="5")
        
        self.passwordField = Entry(self.labelframe1)
        self.passwordField['textvariable'] = self.default_password
        self.passwordField["show"] = u'●' #設定密碼掩蓋符號
        self.passwordField["width"] = 20
        self.passwordField.grid(row=1, column=1, columnspan=2, sticky=W, pady="5")


        self.SMTPText = Label(self.labelframe1)
        self.SMTPText["text"] = u"SMTP伺服器："
        self.SMTPText.grid(row=2, column=0, sticky=E, pady="5")
        self.SMTPList_value = StringVar()
        self.SMTPList = ttk.Combobox(self.labelframe1, textvariable=self.SMTPList_value)
        self.SMTPList['values'] = (u'清華大學SMTP', 'Gmail SMTP')
        self.SMTPList['state'] = 'readonly'
        self.SMTPList.grid(row=2, column=1, sticky=W, pady="5") 

        self.testmailText = Label(self.labelframe1)
        self.testmailText["text"] = u"測試信箱："
        self.testmailText.grid(row=4, column=0, sticky=E, pady="5")
        self.testmailField = Entry(self.labelframe1)
        self.testmailField['textvariable'] = self.default_testmail
        self.testmailField["width"] = 40
        self.testmailField.grid(row=4, column=1, columnspan=2, sticky=W, pady="5")

        self.test = Button(self.labelframe1)
        self.test["text"] = u"寄信測試信"
        self.test.grid(row=5, column=0, columnspan=2,sticky=W, pady="5")
        self.test["command"] = self.clicktest

        self.checkboxvar = IntVar()
        self.saveinfo = Checkbutton(self.labelframe1, text="儲存設定", variable=self.checkboxvar, command = self.saveconfig)
        self.saveinfo.grid(row=5,column=1, columnspan=2,sticky=E, pady="5")


        #框架2
        
        self.labelframe2 = LabelFrame(self, text="其他設定")
        self.labelframe2.config(width=200)
        self.labelframe2.grid(row=1, column=0, columnspan=4, padx="5" ,sticky=W)

        self.default_year = StringVar()
        self.default_semester = StringVar()
        self.default_course = StringVar()
        self.default_exam = StringVar()

        self.yearText = Label(self.labelframe2)
        self.yearText["text"] = u"　 　學年度："
        self.yearText.grid(row=0, column=0, sticky=E, pady="5")
        self.yearField = Entry(self.labelframe2)
        self.yearField["textvariable"] = self.default_year
        self.yearField["width"] = 5
        self.yearField["state"] = 'readonly'
        self.yearField.grid(row=0, column=1, columnspan=2, sticky=W, pady="5")

        self.semesterText = Label(self.labelframe2)
        self.semesterText["text"] = u"　　　學期："
        self.semesterText.grid(row=1, column=0, sticky=E, pady="5")
        self.semesterField = Entry(self.labelframe2)
        self.semesterField["textvariable"] = self.default_semester
        self.semesterField["width"] = 5
        self.semesterField["state"] = 'readonly'
        self.semesterField.grid(row=1, column=1, columnspan=2, sticky=W, pady="5")

        self.courseText = Label(self.labelframe2)
        self.courseText["text"] = u"　課程名稱："
        self.courseText.grid(row=2, column=0, sticky=E, pady="5")
        self.courseField = Entry(self.labelframe2)
        self.courseField["textvariable"] = self.default_course
        self.courseField["width"] = 40
        self.courseField["state"] = 'readonly'
        self.courseField.grid(row=2, column=1, columnspan=2, sticky=W, pady="5")

        self.examText = Label(self.labelframe2)
        self.examText["text"] = u"考試項目："
        self.examText.grid(row=3, column=0, sticky=E, pady="5")
        self.examField = Entry(self.labelframe2)
        self.examField["textvariable"] = self.default_exam
        self.examField["width"] = 30
        self.examField["state"] = 'readonly'
        self.examField.grid(row=3, column=1, columnspan=2, sticky=W, pady="5")

        self.send = Button(self.labelframe2)
        self.send["text"] = u"確認"
        self.send.grid(row=4, column=0, columnspan=2, pady="5")
        self.send["command"] = self.preview

        self.send = Button(self.labelframe2)
        self.send["text"] = u"重新載入"
        self.send.grid(row=4, column=1, columnspan=2, pady="5")
        self.send["command"] = self.loadexcel

        self.send = Button(self.labelframe2)
        self.send["text"] = u"測試鍵"
        #self.send.grid(row=4, column=2, sticky=W, pady="5")
        self.send["command"] = self.click_test      

        self.StudentCountText = Label(self)
        self.StudentCountText["text"] = u"學生人數："
        self.StudentCountText.grid(row=2, column=0, sticky=W, pady="5")
         
        self.StatusCountText = Label(self)
        self.StatusCountText["text"] = u"目前狀態：待命中"
        self.StatusCountText.grid(row=3, column=0, sticky=W)

        
global root
root = Tk()
root.title(u"EDIC_LAB 成績通知系統V" + version)
root.resizable(0,0)
app = GUI_Window(master=root)

app.mainloop()
