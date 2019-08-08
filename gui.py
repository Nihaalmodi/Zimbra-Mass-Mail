import wx
import os
import logging
import pandas as pd
from utils import send_mail
import json
import re
logger = logging.getLogger('AutomatedMail')

class SettingsFrame(wx.Frame):
    def __init__(self, parent,position):
        super().__init__(parent=parent,title='Automated Mailer Settings')
        self.SetPosition(position)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        mail_type_sizer = wx.BoxSizer(wx.HORIZONTAL)
        mail_type_sizer.Add(wx.StaticText(self,-1,"Email Domain: "), 0, wx.ALL | wx.EXPAND, 5)
        self.mail_type = wx.ComboBox(self, size=wx.DefaultSize, choices=['IITKGPMAIL (Zimbra)'],style=wx.CB_READONLY)
        mail_type_sizer.Add(self.mail_type, 0, wx.ALL | wx.EXPAND, 5)

        email_sizer = wx.BoxSizer(wx.HORIZONTAL)
        email_sizer.Add(wx.StaticText(self,-1,"Email ID: "), 0, wx.ALL | wx.EXPAND, 5)
        self.email = wx.TextCtrl(self, size=(200,-1))
        email_sizer.Add(self.email, 0, wx.ALL | wx.EXPAND, 5)
        
        password_sizer = wx.BoxSizer(wx.HORIZONTAL)
        password_sizer.Add(wx.StaticText(self,-1,"Password: "), 0, wx.ALL | wx.EXPAND, 5)
        self.password = wx.TextCtrl(self, size=(200,-1),style = wx.TE_PASSWORD)
        password_sizer.Add(self.password, 0, wx.ALL | wx.EXPAND, 5)

        save_btn = wx.Button(self, label='Save')
        save_btn.Bind(wx.EVT_BUTTON, self.save)


        sizer.Add(mail_type_sizer, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(email_sizer, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(password_sizer, 0, wx.ALL| wx.CENTER , 5)
        sizer.Add(save_btn, 0, wx.ALL | wx.CENTER, 5)
        
        self.SetBackgroundColour("#666")
        self.SetSizer(sizer)
        if os.path.exists("credentials.json"):
            file = open("credentials.json","r")
            data = json.load(file)
            file.close()
            self.mail_type.SetValue(data.get("type"))
            self.email.SetValue(data.get("email"))
            self.password.SetValue(data.get("password"))

        self.Show()
    
    def save(self,event):
        email = str(self.email.GetValue())
        match = re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', email)
        if match == None:
            wx.MessageBox("Bad Email", 'Info', wx.OK)
            return False
        #TODO: Apply minimal encryption
        password = self.password.GetValue()
        file = open("credentials.json","w")
        json.dump({"type":str(self.mail_type.GetValue()),"email":email,"password":password},file)
        file.close()
        self.Destroy()


class MainFrame(wx.Frame):    
    def __init__(self,):
        super().__init__(parent=None, title='Automated Mailer')
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        PreferencesMenu = wx.Menu()
        HelpMenu = wx.Menu()
        fileItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit application')
        connect_email = PreferencesMenu.Append(wx.ID_ANY, 'Connect Email','')
        Help = HelpMenu.Append(wx.ID_ANY, 'Usage', '')
        menubar.Append(fileMenu, '&File')
        menubar.Append(PreferencesMenu, '&Preferences')
        menubar.Append(HelpMenu, '&Help')
        
        self.SetMenuBar(menubar)
        self.Bind(wx.EVT_MENU, self.OnQuit, fileItem)
        self.Bind(wx.EVT_MENU, self.connect_email, connect_email)
        self.Bind(wx.EVT_MENU, self.display_help, Help)
        
        sizer = wx.BoxSizer(wx.VERTICAL)  

        excel_sizer = wx.BoxSizer(wx.HORIZONTAL)
        excel_sizer.Add(wx.StaticText(self,-1,"Excel Sheet Path:"), 0, wx.ALL | wx.EXPAND, 5)
        self.path = wx.TextCtrl(self, size=(200,-1))
        excel_sizer.Add(self.path, 0, wx.ALL | wx.EXPAND, 5)

        browse_btn = wx.Button(self, label='Browse')
        browse_btn.Bind(wx.EVT_BUTTON, self.browse_excel)
        excel_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(excel_sizer,0, wx.ALL | wx.EXPAND, 5)

        attachment_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        attachment_sizer.Add(wx.StaticText(self,-1,"Attachments:"), 0, wx.ALL | wx.EXPAND, 5)
        self.attachments = wx.TextCtrl(self,size=(200,-1))
        attachment_sizer.Add(self.attachments, 0, wx.ALL | wx.EXPAND, 5)
        
        browse_btn = wx.Button(self, label='Browse')
        browse_btn.Bind(wx.EVT_BUTTON, self.browse_attachments)
        attachment_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(attachment_sizer,0, wx.ALL | wx.EXPAND, 5)

        
        subject_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        subject_sizer.Add(wx.StaticText(self,-1,"Subject:"), 0, wx.ALL | wx.EXPAND, 5)
        self.subject = wx.TextCtrl(self,style =  wx.SUNKEN_BORDER  | wx.EXPAND, size=(200,-1))
        subject_sizer.Add(self.subject, 0,wx.ALL | wx.EXPAND, 5)
        sizer.Add(subject_sizer,0, wx.ALL | wx.EXPAND, 5)

        sizer.Add(wx.StaticText(self,-1,"Body:"), 0, wx.ALL, 5)
        self.body = wx.TextCtrl(self, style =  wx.TE_MULTILINE | wx.SUNKEN_BORDER | wx.EXPAND,size=(-1,200))
        sizer.Add(self.body, 0, wx.ALL | wx.EXPAND, 5)

        self.send_btn = wx.Button(self, label='Send Mail(s)')
        self.send_btn.Bind(wx.EVT_BUTTON, self.send)
        sizer.Add(self.send_btn, 0, wx.ALL | wx.CENTER, 5)

        self.statusbar = self.CreateStatusBar(1)
        self.statusbar.SetStatusText('Ready!')
        
        self.SetBackgroundColour("#666")
        self.SetSizer(sizer)
        self.Show()
    
    def OnQuit(self,event):
        self.Destroy()

    def connect_email(self,event):
        SettingsFrame(self,self.GetPosition())
        
    def display_help(self,event):
        #TODO: display help instructions
        pass

    def browse_excel(self,event):
        dlg = wx.FileDialog(
            self, message="Choose a file",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard="All files (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_CHANGE_DIR
            )
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()[0]
            if(".xlsx" !=paths.split(".")[-1]):
                self.statusbar.SetStatusText('The selected file may not be supported!!')
            self.path.SetValue(paths)
        dlg.Destroy()

    def browse_attachments(self,event):
        dlg = wx.FileDialog(
            self, message="Choose a file",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard="All files (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_CHANGE_DIR | wx.FD_MULTIPLE
            )
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            self.attachments.SetValue(",".join(paths))
        dlg.Destroy()

    def send(self,event):
        self.send_btn.Disable()
        if len(self.path.GetValue())==0:
            self.statusbar.SetStatusText("Please Enter Excel File Path")
            return False
        try:
            df = pd.read_excel(self.path.GetValue())
        except FileNotFoundError:
            self.statusbar.SetStatusText("Excel File not found")
        body = self.body.GetValue()
        file = open("credentials.json","r")
        data = json.load(file)
        file.close()
        log = []
        for _, rows in df.iterrows():
            self.statusbar.SetStatusText('Sending mail to: {}'.format(rows[0]))
            header = df.columns
            dictionary = {header[i]:rows[i] for i in range(1,len(rows))}
            body = body.format(**dictionary)
            subject = self.subject.GetValue()
            attachments = None
            if len(self.attachments.GetValue())>0:
                attachments = self.attachments.GetValue()
            response = send_mail(data.get("email"),data.get("password"), rows[0], subject, body, file=attachments,SMTPserver=data.get("type"))
            if response==200:
                self.statusbar.SetStatusText('Mail Sent to: {}'.format(rows[0]))
                log.append((rows[0],"sent"))
            else:
                self.statusbar.SetStatusText('Mail Sent failed: {}'.format(rows[0]))
                log.append((rows[0],"failed"))
        message = "Mail Status\n"
        for mail,status in log:
            message += mail + "--" + status +"\n"
        wx.MessageBox(message, 'Info', wx.OK)
        self.send_btn.Enable()


# class LoginPanel(wx.Panel):
#     def __init__(self, parent):
#         self.parent = parent
#         wx.Panel.__init__(self, parent)
#         sizer = wx.BoxSizer(wx.VERTICAL)  

#         password_text = wx.StaticText(self,-1,"Password:")
#         sizer.Add(password_text, 0, wx.ALL | wx.EXPAND, 5)
#         self.password = wx.TextCtrl(self,style = wx.TE_PASSWORD)
#         sizer.Add(self.password, 0, wx.ALL | wx.EXPAND, 5)
        
#         login_btn = wx.Button(self, label='Login')
#         login_btn.Bind(wx.EVT_BUTTON, self.login)
#         sizer.Add(login_btn, 0, wx.ALL | wx.CENTER, 5)
#         self.SetSizer(sizer)
    
#     def login(self, event):
#         if not self.validate_password():
#             self.statusbar.SetStatusText("Error!! Wrong Password")
#             logger.error("Error wrong password")
#             return False
#         self.Destroy()
#         self.parent.Destroy()
#         MainFrame(self.parent.GetPosition())
    
#     def validate_password(self):
#         return True
        
# class LoginFrame(wx.Frame):    
#     def __init__(self):
#         super().__init__(parent=None, title='Automated Mailer')
#         login = LoginPanel(self)

#         self.statusbar = self.CreateStatusBar(1)
#         self.statusbar.SetStatusText('Please Enter Password to Login')
#         self.Show()
        
 

#     # def create_main_panel(self):
#     #     self.main_panel.Show()
        
    
    

if __name__ == '__main__':
    app = wx.App()
    frame = MainFrame()
    app.MainLoop()