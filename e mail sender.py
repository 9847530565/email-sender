from tkinter import *
from tkinter.filedialog import askopenfile
from tkinter import simpledialog
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
screen=Tk()
attachoption=False
width,height=500,500
screen.geometry(f"{width}x{height}")
photo=PhotoImage(file="files//Settings.png")
screen.iconphoto(True,photo)
screen.title("E mail sender")
screen.resizable(0,0)
canvas=Canvas(screen,width=width,height=height)
canvas.pack()
image1=PhotoImage(file="files//background.png")
canvas.create_image(0,0,anchor=NW,image=image1)
def openfile():
    global file
    file=askopenfile(mode="r",filetypes=[("Excel Work Sheets","*.xlsx")])
    if file is not None:
        identity.set(file.name)
def addfiles():
    global attachoption
    attachoption=True
    filename=askopenfile(mode="a",filetypes=[("Excel Work Sheets",".xlsx"),("PDF Files","*.pdf"),("PNG image","*.png"),("JPEG image","*.jpeg"),("JPG images","*.jpg")])
    x=filename.name
    filename=open(x,"rb")
    global msg
    msg=MIMEMultipart()
    msg["From"]=usermail
    msg["To"]=identity.get()
    body=box2.get("1.0","end")
    box2.delete("1.0","end")
    msg["Subject"]="E mail sending application"
    msg.attach(MIMEText(body, 'plain'))
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((filename).read())
    encoders.encode_base64(p)
    #msg.attach(MIMEText(box2.get("1.0","end"), 'plain'))
    p.add_header('Content-Disposition', "attachment; filename= %s" % x)
    msg.attach(p)
    if filename is not None:
        #box2.delete("1.0","end")
        box2.insert("1.0",x)
def resetentries():
    option.set(0)
    identity.set("")
    box2.delete("1.0","end")
def sendemail():
    id1=identity.get()
    message=box2.get("1.0","end")
    if id1 is not None:
        mailtype=option.get()
        server=smtplib.SMTP("smtp.gmail.com",587)
        server.ehlo()
        server.starttls()
        server.login(usermail,userpassword)
        if mailtype==1:
            server.sendmail(usermail,id1,msg.as_string())
            atttachoption=False
        elif mailtype==2:
            id1.replace('/', "\\")
            df=pd.read_excel(id1)
            df1=df["email"]
            for i in df1:
                server.sendmail(usermail,i,msg.as_string())
        server.quit()
    else:
        messagebox.showerror("Fatal Error","Please enter email address")
global option
option=IntVar()
option.set(0)
Label(screen,text="EMAIL SENDER",bg="green",fg="cyan",font=("Calibri",40)).place(relx=0.2,x=-2,y=10)
Radiobutton(screen,text="Single",variable=option,value=1,bg="green",font=("Calibri",10),fg="cyan").place(relx=0.1,x=-2,y=140)
Radiobutton(screen,text="Multiple",variable=option,value=2,bg="green",font=("Calibri",10),fg="cyan").place(relx=0.3,x=-2,y=140)
Label(screen,text="Enter email id",bg="green",font=("Calibri",10),fg="cyan").place(relx=0,x=-2,y=2*height//5)
global identity,box2
identity=StringVar()
Entry(screen,text=identity,width=30,bg="red",fg="cyan").place(relx=0.25,x=-2,y=2*height//5)
Button(screen,text="browse",command=openfile).place(relx=0.7,x=-2,y=2*height//5)
Button(screen,text="add file",command=addfiles).place(relx=0.7,x=-2,y=3*height//5)
Label(screen,text="enter content",bg="green",font=("Calibri",10),fg="cyan").place(relx=0,x=-2,y=height//2)
box2=Text(screen,height=10,width=30,bg="red",fg="cyan",font=("Calibri",10))
box2.place(relx=0.25,x=-2,y=height//2)
Button(screen,text="Send",command=sendemail).place(relx=0.5,x=-2,y=4.5*height//5)
Button(screen,text="reset",command=resetentries).place(relx=0.7,x=-2,y=4.5*height//5)
global username,userpassword
usermail=simpledialog.askstring("input","enter your unofficial email id",parent=screen)
userpassword=simpledialog.askstring("input","enter your password",parent=screen)
if usermail==None or userpassword==None:
    screen.destroy()
