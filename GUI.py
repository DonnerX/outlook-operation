#coding:utf-8

import win32com.client
import Tkinter as tk
import re

class outlook(object):
    def __init__(self):
        self.outlook=win32com.client.gencache.EnsureDispatch("Outlook.Application")
        self.outlook_namespace=self.outlook.GetNamespace("MAPI")
        self.inbox=self.outlook_namespace.Folders["support.cn@etas.com"].Folders['1 MCD'].Folders['2_MCD_Done'].Folders['test']
        self.messages=self.inbox.Items
        
        
        self.mail=self.outlook.CreateItem(win32com.client.constants.olMailItem)

    def getaddress(self,dstart,dend):
        l=[]
        for t in range(self.messages.Count):
            item=self.messages.Item(t+1)
            date=item.CreationTime
            if(len(str(date.day))==1):
                day='0'+str(date.day)
            else:
                day=str(date.day)
            datestr=str(date.year)+str(date.month)+day
            if datestr<dstart:
                break
            if (datestr>dstart)&(datestr<dend):
                try:
                    name=item.SenderEmailAddress
                    l.append(name)
                except:
                    print 'one Email has no Address'
        return l
    def fil(self,l):
        f=[]
        r=r'@'
        r2=r'ETAS|bosch'
        for add in l:
            if (re.search(r,add)!=None)&(re.search(r2,add)==None):
                f.append(add)
        return f
    def write_to_txt(self,l):
        with open('address.txt','w') as f:
            for s in l:
                f.writelines(s+'\n')
            f.close()

    def delete_same(self,l):
        l.sort()
        for s in l:
            n=l.count(s)
            if n>1:
                for i in range(n-1):
                    l.remove(s)
        return l
    def send(self,l,survey,sub):
        self.mail.Body=survey
        self.mail.Subject=sub
        quo=';'
        s=quo.join(l)
        print s
        #self.mail.BCC=(s)
        #self.mail.Send()
        #print 'send'
                
        
#---------------------------------------------------------------------------------#            

def send():
    dstart=estart.get()
    dend=eend.get()
    ou=outlook()
    l=ou.getaddress(dstart,dend)
    f=ou.fil(l)
    de=ou.delete_same(f)
    ou.write_to_txt(de)

    survey=esurvey.get()
    sub=esub.get()
    ou.send(de,survey,sub)
#---------------------------------------------------------------------------------#
window=tk.Tk()
lstart=tk.Label(text='start')
lstart.pack()
estart=tk.Entry()
estart.pack()

lend=tk.Label(text='end')
lend.pack()
eend=tk.Entry()
eend.pack()

lsurvey=tk.Label(text='问卷网址')
lsurvey.pack()
esurvey=tk.Entry()
esurvey.pack()

lsub=tk.Label(text='邮件标题')
lsub.pack()
esub=tk.Entry()
esub.pack()

b=tk.Button(text='send',command=send)
b.pack()


window.mainloop()
    
