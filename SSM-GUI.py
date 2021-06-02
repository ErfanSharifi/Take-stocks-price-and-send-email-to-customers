# -*- coding: utf-8 -*-
#!/usr/bin/python

from tkinter import *
from tkinter import ttk
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from unidecode import unidecode
import smtplib
import Check_Email as chem
import Make_Excel as mex
import Send_mail as sm
import Send_Attach_mail as sam
import csv
import xlsxwriter

class FirstGUI:

    
    def __init__(self,root):

        self.list2 = []


        #########################################################################################################
        ##################################                نام سهم                ###############################
        #########################################################################################################
        
        self.lbl_sharename = Label(root, text= "نام سهم:")
        self.lbl_sharename.grid(row = 2, column = 0,padx =10, pady =10, sticky = W)

        self.text_sharename = Text(root,width=30,height=1)
        self.text_sharename.grid(row=2,column=1)
        
        #########################################################################################################
        ##################################                تعداد سهم                #############################
        #########################################################################################################

        self.lbl_numbershare = Label(root, text= "تعداد سهم:").grid(row = 2, column = 2,padx =10, pady =10, sticky = W)
        self.text_numbershare = Text(root,width=30,height=1)
        self.text_numbershare.grid(row=2,column=3)
        self.text_numbershare.get("1.0", END)




        #########################################################################################################
        ##################################                محاسبه                   ##############################
        #########################################################################################################

        self.lbl_between= Label(root, text = "محاسبه:").grid(row= 4, column =2, padx =10, pady = 10, sticky = W)
        self.btn_calculate = Button(root, text = ("محاسبه"), width = 34, height = 1, command = self.TakeShareName)
        self.btn_calculate.grid(row= 4, column = 3)

        #########################################################################################################
        ##################################                ارزش سهم                 #############################
        #########################################################################################################

        self.lbl_shareday = Label(root, text= "ارزش سهم:").grid(row = 3, column = 0,padx =10, pady =10, sticky = W)
        self.text_shareday = Text(root,width=30,height=1)
        self.text_shareday.config(state = DISABLED)
        self.text_shareday.grid(row=3,column=1)

        #########################################################################################################
        ##################################                ارزش کل سهم                ###########################
        #########################################################################################################

        self.lbl_total = Label(root, text= "ارزش کل سهم:").grid(row = 3, column = 2,padx =10, pady =10, sticky = W)
        self.text_total = Text(root,width=30,height=1)
        self.text_total.config(state = DISABLED)
        self.text_total.grid(row=3,column=3)

        #########################################################################################################
        ##################################                قیمت روز سهم                ##########################
        #########################################################################################################

        self.lbl_dayprice = Label(root, text= "قیمت روز سهم:").grid(row = 4, column = 0,padx =10, pady =10, sticky = W)
        self.text_dayprice = Text(root,width=30,height=1)
        self.text_dayprice.config(state = DISABLED)
        self.text_dayprice.grid(row=4,column=1)

        #########################################################################################################
        ##################################                نام فایل اکسل را انتخاب کنید   ######################
        #########################################################################################################

        self.lbl_nameexcel = Label(root, text= "نام فایل اکسل را انتخاب کنید:")
        self.lbl_nameexcel.grid(row = 5, column = 0,padx =10, pady =10, sticky = W)

        self.text_nameexcel = Text(root,width=30,height=1)
        self.text_nameexcel.grid(row=5,column=1)

        #########################################################################################################
        ####################        برای ساخت فایل اکسل دکمه را فشار دهید     ################################
        #########################################################################################################

        self.lbl_makeexcell= Label(root,text = "برای ساخت فایل اکسل دکمه را فشار دهید:")
        self.lbl_makeexcell.grid(row= 5, column =2,padx =10, pady =10, sticky = W)
        
        self.btn_makeexcell = Button(root, text = ("Excel"), width = 34, height = 1, command = self.MAkeExcel)
        self.btn_makeexcell.grid(row= 5, column = 3)

        #########################################################################################################
        ##################################                Email   and   sms   ###################################
        #########################################################################################################

        def callback(stat):
            w = self.cmb_email_sms.get()
            if w == "Email":
                self.text_email.config(state = NORMAL)
            else:
                self.text_sms.config(state = NORMAL)
                

        self.lbl_Email= Label(root,text = "لطفا یک گزینه را انتخاب کنید:")
        self.lbl_Email.grid(row= 6, column =0,padx =10, pady =10, sticky = W)
        self.stat = StringVar()
        self.cmb_email_sms = ttk.Combobox(root, textvariable= self.stat, values=("Email"))#,"SMS"
        self.cmb_email_sms.grid(row = 6, column = 1)
        self.cmb_email_sms.bind("<<ComboboxSelected>>", callback)


        self.lbl_email = Label(root, text = "Email:").grid(row = 7, column = 0, padx =10, pady = 10, sticky =NSEW)
        self.text_email = Text(root,width=30,height=1)
        self.text_email.config(state = DISABLED)
        self.text_email.grid(row=7,column=1)
        
        
        self.lbl_sms = Label(root, text = "SMS:").grid(row = 7, column = 2, padx =10, pady = 10, sticky =NSEW)
        self.text_sms = Text(root,width=30,height=1)
        self.text_sms.config(state = DISABLED)
        self.text_sms.grid(row=7,column=3)

        #########################################################################################################
        ##################################                Send and Clear      ###################################
        #########################################################################################################

        self.btn_send = Button(root, text = ("Send"), width = 34, height = 1, command = self.SendMai)
        self.btn_send.grid(row= 8 ,column = 0, padx =40, pady =10)

        self.btn_clear = Button(root, text = ("Clear"), width = 34, height = 1, command = self.Clear)
        self.btn_clear.grid(row= 8 ,column = 3, padx =10, pady =10)

        #########################################################################################################
        ##################################                ERRORS Labels       ###################################
        #########################################################################################################

        self.text_errors = Text(root,width=42,height=4)
        self.text_errors.config(state = DISABLED)
        self.text_errors.grid(row=9,column=0,columnspan = 8, padx = 30, pady = 10, sticky = NSEW)

        #########################################################################################################
        ##################################                TakeShareName       ###################################
        #########################################################################################################

    def TakeShareName(self):
        
        self.share_list = []
        
        print (self.share_list)

        val = True
        check = self.text_numbershare.get("1.0", END)
        b = ['1','2','3','4','5','6','7','8','9','0']
        c=list(check.strip())
        for e in c:
            if e not in b: 
                val =False
                break
        if val:
            global q, name
            name = self.text_sharename.get("1.0",END)
            options = Options()
            options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            driver = webdriver.Chrome(chrome_options=options, executable_path= r'C:\chromedriver')
            driver.get('http://www.fipiran.com/Symbol?symbolpara=' + name)
            soup = BeautifulSoup(driver.page_source,"lxml")
            item = soup.find('span', {'id' : 'PDrCotVal'})
            item = item.get_text()
            q = item
            
            self.text_dayprice.config(state = NORMAL)
            self.text_dayprice.insert(END, q)
            self.text_dayprice.config(state = DISABLED)

            q = unidecode(q)
            q = q.replace(",", "")
            q = (int(q))
            lk = self.text_numbershare.get("1.0", END)
            lk = (int(lk))
            su = (q*lk)

            self.text_shareday.config(state = NORMAL)
            self.text_shareday.insert(END, su)
            self.text_shareday.config(state = DISABLED)


        #########################################################################################################
        #########################################################################################################


            n = self.text_sharename.get("1.0",END)

            k = self.text_numbershare.get("1.0", END)

            self.text_shareday.config(state = NORMAL)
            a = self.text_shareday.get("1.0", END)
            self.text_shareday.config(state = DISABLED)

            self.text_dayprice.config(state = NORMAL)
            m = self.text_dayprice.get("1.0", END)
            self.text_dayprice.config(state = DISABLED)

            self.share_list.append(n)
            self.share_list.append(k)
            self.share_list.append(m)
            self.share_list.append(a)

            stri = str(self.share_list)

            stri = stri.replace("]", '')
            stri = stri.replace("'", '')
            stri = stri.replace("[", '')
            stri = stri.replace(",", '')
            stri = stri.replace("\\n", '')
            stri = stri.split()


        #########################################################################################################
        #########################################################################################################


        else:
            self.text_errors.config(state = NORMAL)
            self.text_errors.insert(END, " لطفا عدد را به درستی وارد کنین")
            self.text_errors.config(state = DISABLED)  

        
        self.list2.append(stri)
        print (self.list2)





        #########################################################################################################
        ##################################                CheckMail           ###################################
        #########################################################################################################  



    # def CheckMail(self):
    #     mail = self.text_email.get("1.0",END)
    #     obj = chem.Mail()
    #     obj.CheckMail(mail)
    #     k = obj.CheckMail(mail)
    #     if (k != True):
    #         self.text_errors.config(state = NORMAL)
    #         self.text_errors.insert(END,"لطفا ایمیل را به درستی وارد کنید...")
    #         self.text_errors.config(state = DISABLED)
        


        #########################################################################################################
        ##################################                Make Excel          ###################################
        #########################################################################################################



    def MAkeExcel(self):


        name = self.text_nameexcel.get("1.0", END)
        name = str(name)
        name = name.replace("\n", '')

        obj = mex.Excel()
        obj.Make(name,self.list2)


        #########################################################################################################
        ##################################               Send Mail            ###################################
        #########################################################################################################
   

    def SendMai(self):
        maill = self.text_email.get("1.0",END)
        obj = chem.Mail()
        obj.CheckMail(maill)
        k = obj.CheckMail(maill)
        if (k != True):
            self.text_errors.config(state = NORMAL)
            self.text_errors.insert(END,"لطفا ایمیل را به درستی وارد کنید...")
            self.text_errors.config(state = DISABLED)
        else:
            obj = sam.Send()
            obj.GiveMail(maill)
        

        #########################################################################################################
        ##################################                Def Clear           ###################################
        #########################################################################################################


    def Clear(self):

        self.text_sharename.delete("1.0",END)
        self.text_numbershare.delete("1.0",END)

        self.text_shareday.config(state = NORMAL)
        self.text_shareday.delete("1.0",END)
        self.text_shareday.config(state = DISABLED)

        self.text_total.config(state = NORMAL)
        self.text_total.delete("1.0",END)
        self.text_total.config(state = DISABLED)

        self.text_dayprice.config(state = NORMAL)
        self.text_dayprice.delete("1.0",END)
        self.text_dayprice.config(state = DISABLED)

        self.text_nameexcel.delete("1.0",END)
        self.text_email.delete("1.0",END)
        self.text_sms.delete("1.0",END)
        self.text_errors.delete("1.0",END)
        

        #########################################################################################################
        ##################################                Main                ###################################
        #########################################################################################################


def main():
    root =Tk()
    root.geometry("1070x400+500+300")
    root.title("                                                                                                                                                به برنامه محاسبه سهام خوش آمدید ")
    root.resizable(0,0)
    obj =FirstGUI(root)
    root.mainloop()
main()