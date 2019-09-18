from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import smtplib, ssl
import xlrd  
import time
import webbrowser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from os.path import basename
import email.mime.application
from email import encoders
from validate_email import validate_email

file =''

def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askopenfilename()
    file_path.set(filename)
    global filee
    filee= filename
    print(filee)

file_attach=''
def browse_attach():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_attach
    attach_name = filedialog.askopenfilename()
    file_path.set(attach_name)
    global file_attach
    file_attach= attach_name
    print(attach_name)
    

def sendemail():
    try:
        port = 587  # For starttls
        smtp_server = "smtp.gmail.com"
        
        sender_email = account.get()
        pasword = password.get()
        b=start.get()
        s=end.get()
        c=sutun.get()
        mail_content = msgbody.get('1.0','end')

        message = MIMEMultipart()
        message['From'] = account.get()
        
        message['Subject'] = subject.get()
        print(file)
        message.attach(MIMEText(mail_content, 'plain'))
        if file_attach!='':
            filename=file_attach
            fo=open(filename,'rb')
            print(filename.split('/',1))
            attach = email.mime.application.MIMEApplication(fo.read(),_subtype=filename.split('/',1))
            fo.close()
            attach.add_header('Content-Disposition','attachment',filename=file_attach.split('/')[-1])
            message.attach(attach)
        
  
        loc = filee
        wb = xlrd.open_workbook(loc) 
        sheet = wb.sheet_by_index(0) 
        
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, port) as server:
            server.ehlo()  # Can be omitted
            server.starttls(context=context)
            server.ehlo()  # Can be omitted
            server.login(sender_email, pasword)
            for i in range(b,s):
                text = message.as_string()
                if  (validate_email(sheet.cell_value(i, c))):
                    #message['To'] = sheet.cell_value(i, c)
                    server.sendmail(sender_email,sheet.cell_value(i, c) , text)
                    print(i,sheet.cell_value(i, c),': ugradyldy')
                else:
                    print(i,sheet.cell_value(i, c),': yalnysh email address')
                if i%10==0:
                    print("I am sleeping")
                    time.sleep(10)

            
        ttk.Label(mainframe, text="Emaillerin ahlisi ugradyldy").grid(column=4,row=12,sticky=W)

    except Exception as e:
        ttk.Label(mainframe, text=str(e)).grid(column=4,row=9,sticky=W)

def setup(event):
    webbrowser.open_new(r"https://www.google.com/settings/security/lesssecureapps")
    
        
root = Tk()
root.title("GMail ulanyp Maillere hat ugratmak")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

account = StringVar()
password = StringVar()
#receiver = StringVar()
subject = StringVar()
msgbody = StringVar()
start=IntVar()
end=IntVar()
sutun=IntVar()
file_path=StringVar()
attach_path=StringVar()

a = Label(mainframe, text="Gmail akandynyzyn kabir nastroykalaryny duzedin", fg="blue", cursor="hand2")
a.grid(columnspan=2,column=3, row=0, sticky=W)
a.bind("<Button-1>", setup)


ttk.Label(mainframe, text="Email adresiniz: ").grid(column=0, row=1, sticky=W)
account_entry = ttk.Entry(mainframe, width=30, textvariable=account)
account_entry.grid(column=4, row=1, sticky=(W, E))

ttk.Label(mainframe, text="Parolynyz: ").grid(column=0, row=2, sticky=W)
password_entry = ttk.Entry(mainframe, show="*", width=30, textvariable=password)
password_entry.grid(column=4, row=2, sticky=(W, E))

ttk.Label(mainframe, text="Tema: ").grid(column=0, row=6, sticky=W)
subject_entry = ttk.Entry(mainframe, width=30, textvariable=subject)
subject_entry.grid(column=4, row=6, sticky=(W, E))

ttk.Label(mainframe, text="Hatyn Mazmuny: ").grid(column=0, row=7, sticky=W)
msgbody = Text(mainframe, width=30, height=10)
msgbody.grid(column=4, row=7, sticky=(W, E))


ttk.Label(mainframe, text="Basy: ").grid(column=0, row=8, sticky=W)
start_entry = ttk.Entry(mainframe, width=15, textvariable=start)
start_entry.grid(column=4, row=8, sticky=(W, E))

ttk.Label(mainframe, text="Sony: ").grid(column=0, row=9, sticky=W)
end_entry = ttk.Entry(mainframe, width=15, textvariable=end)
end_entry.grid(column=4, row=9, sticky=(W, E))

ttk.Label(mainframe, text="Sutuni(A=0,B=1,C=2...): ").grid(column=0, row=10, sticky=W)
end_entry = ttk.Entry(mainframe, width=15, textvariable=sutun)
end_entry.grid(column=4, row=10, sticky=(W, E))

lbl1 = ttk.Label(mainframe,textvariable=file_path)
lbl1.grid(column=4,row=11,sticky=E)

button2 = ttk.Button(mainframe,text="Haysy fayldan ugratmaly", command=browse_button)
button2.grid(column=0,row=11,sticky=(W,E))

lbl2 = ttk.Label(mainframe,textvariable=attach_path)
lbl2.grid(column=4,row=12,sticky=E)

button3 = ttk.Button(mainframe,text="Ugratjak faylynyzy saylan", command=browse_attach)
button3.grid(column=0,row=12,sticky=(W,E))



ttk.Button(mainframe, text="Emailleri ugrat", command=sendemail).grid(column=4,row=13,sticky=(W,E))

for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

account_entry.focus()

root.mainloop()
