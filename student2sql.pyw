import os
from tkinter import *
import xlsxwriter
import tkinter as tk
from tkinter import ttk
import mysql.connector as mc
from tkinter import messagebox
import tkinter.messagebox as tm

temp1="hii5.xlsx"
cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
file = open(temp1, "w+")
workbook = xlsxwriter.Workbook(temp1)

def exportxl():
    worksheet = workbook.add_worksheet() 
    worksheet.write('A1', 'Roll NO') 
    worksheet.write('B1', 'First Name') 
    worksheet.write('C1', 'Last Name') 
    worksheet.write('D1', 'Mobile No')
    worksheet.write('E1', 'Subject 1')
    worksheet.write('F1', 'Subject 2')
    worksheet.write('G1', 'Subject 3')
    worksheet.write('H1', 'Subject 4')
    worksheet.write('I1', 'Total Marks')
    worksheet.write('J1', 'Pass/Fail')
    
    cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
    cur=cn.cursor()
    cur.execute("select * from student2")
    rec=cur.fetchall()[0]
    p=100
    for row in rec:
        p+=1
    row=1
    col=0
    a=[]
    b=[]
    c=[]
    d=[]
    e=[]
    f=[]
    g=[]
    h=[]
    p2=[]
    p3=[]
    for i in range(101, p+1):
        cur.execute("""select roll from student2 where roll =%s""", (i,))
        nm = cur.fetchone()[0]     
        cur.execute("""select first_name from student2 where roll =%s""", (i,))
        fnm = cur.fetchone()[0]
        cur.execute("""select last_name from student2 where roll =%s""",(i,))
        lnm = cur.fetchone()[0]
        cur.execute("""select mobile_no from student2 where roll =%s""", (i,))
        mob = cur.fetchone()[0]
        print(nm)
        a.append(nm)
        b.append(fnm)
        c.append(lnm)
        d.append(mob)
    for item in a:
        print(item)
        worksheet.write(row, col, str(item)) 
        row += 1
    row=1
    for item in b:
        print(item)
        worksheet.write(row, col+1, str(item)) 
        row += 1
    row=1
    for item in c:
        print(item)
        worksheet.write(row, col+2, str(item)) 
        row += 1
    row=1
    for item in d:
        print(item)
        worksheet.write(row, col+3, str(item)) 
        row += 1
    row=1
    
    
    
        
        
    cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
    cur=cn.cursor()
    cur.execute("select * from marks")
    rec1=cur.fetchall()[0]
    
    p1=100
    for row1 in rec1:
        p1+=1
    tmp1=p1+1
    for i in range(101, p1+2):
        cur.execute("""select marks1 from marks where roll =%s""", (i,))
        fnm1 = cur.fetchone()[0]
        cur.execute("""select marks2 from marks where roll =%s""",(i,))
        lnm1 = cur.fetchone()[0]
        cur.execute("""select marks3 from marks where roll =%s""", (i,))
        mob1 = cur.fetchone()[0]
        cur.execute("""select marks4 from marks where roll =%s""",(i,))
        std1 = cur.fetchone()[0]
        gend2=int(fnm1)+int(lnm1)+int(mob1)+int(std1)
        if(gend2<200):
            p3.append("fail")
        else:
            p3.append("pass")
        e.append(fnm1)
        f.append(lnm1)
        g.append(mob1)
        h.append(std1)
        p2.append(gend2)
        print(gend2)
        
    for item in e:
        print(item)
        worksheet.write(row, col+4, str(item)) 
        row += 1
    row=1
    for item in f:
        print(item)
        worksheet.write(row, col+5, str(item)) 
        row += 1
    row=1
    for item in g:
        print(item)
        worksheet.write(row, col+6, str(item)) 
        row += 1
    row=1
    for item in h:
        print(item)
        worksheet.write(row, col+7, str(item)) 
        row += 1
    row=1
    for item in p2:
        print(item)
        worksheet.write(row, col+8, str(item)) 
        row += 1
    row=1
    for item in p3:
        print(item)
        worksheet.write(row, col+9, str(item)) 
        row += 1
    row=1
    workbook.close()
    messagebox.showinfo("Succesful operation", "details exported to excel succesfully!")
    
    
    
    
   
def feestk():
    win2=Toplevel(root)
    rightframe=Frame(win2)
    rightframe.pack(side='left')
    win2.geometry("1350x270+0+170")
    feesframe=Frame(win2)
    feesframe.pack(side='left')
    tree = ttk.Treeview(rightframe)
    tree.pack(padx=5, pady=10)
    roll_no=IntVar()
    marks_1=StringVar()
    marks_2=StringVar()
    marks_3=StringVar()
    marks_4=StringVar()
    l1=Label(feesframe,text="Roll No:")
    l1.grid(row=0,column=0,sticky=W)
    rollentry=Entry(feesframe,width=8,textvariable=roll_no)
    rollentry.grid(row=0,column=1)
    l2=Label(feesframe,text="Subject 1:")
    l2.grid(row=2,column=0,sticky=W)
    m1entry=Entry(feesframe,width=8,textvariable=marks_1)
    m1entry.grid(row=2,column=1)
    l3=Label(feesframe,text="Subject 2:")
    l3.grid(row=3,column=0,sticky=W)
    m2entry=Entry(feesframe,width=8,textvariable=marks_2)
    m2entry.grid(row=3,column=1)
    l4=Label(feesframe,text="Subject 3:")
    l4.grid(row=4,column=0,sticky=W)
    m3entry=Entry(feesframe,width=8,textvariable=marks_3)
    m3entry.grid(row=4,column=1)
    l5=Label(feesframe,text="Subject 4:")
    l5.grid(row=5,column=0,sticky=W)
    m4entry=Entry(feesframe,width=8,textvariable=marks_4)
    m4entry.grid(row=5,column=1)
    rollentry.delete(0,END)
    def put(*args):
        try:
            roll=roll_no.get()
            marks1=marks_1.get()
            marks2=marks_2.get()
            marks3=marks_3.get()
            marks4=marks_4.get()
            cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
            cur=cn.cursor()
            cur.execute("insert into marks(roll,marks1,marks2,marks3,marks4)\
                    values('"+str(roll)+"','"+marks1+"','"+marks2+"','"+marks3+"','"+marks4+"')")
            cur.execute("commit")
            cn.close()
            rollentry.delete(0,END)
            m1entry.delete(0,END)
            m2entry.delete(0,END)
            m3entry.delete(0,END)
            m4entry.delete(0,END)
            
        except:
            messagebox.showinfo("Error X", "Enter All Valid Feilds!")
    def rem(*args):
        try:
            roll=roll_no.get()
            cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
            cur=cn.cursor()
            cur.execute("DELETE FROM marks WHERE roll ='"+str(roll)+"'")
            cur.execute("commit")
            cn.close()
            rollentry.delete(0,END)
            
        except:
            messagebox.showinfo("Error X", "Enter The ROll NO !")
            
    def content(*args):
        
        cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
        cur=cn.cursor()
        cur.execute("select * from marks")
        rec=cur.fetchall()[0]
        p=100
        for row in rec:
            p+=1
        tmp=p+1
    
        for i in tree.get_children():
            tree.delete(i)
        for i in range(101, p+2):
            cur.execute("""select roll from marks where roll =%s""", (i,))
            nm = cur.fetchone()[0]
            cur.execute("""select marks1 from marks where roll =%s""", (i,))
            fnm = cur.fetchone()[0]
            cur.execute("""select marks2 from marks where roll =%s""",(i,))
            lnm = cur.fetchone()[0]
            cur.execute("""select marks3 from marks where roll =%s""", (i,))
            mob = cur.fetchone()[0]
            cur.execute("""select marks4 from marks where roll =%s""",(i,))
            std = cur.fetchone()[0]
            gend=int(fnm)+int(lnm)+int(mob)+int(std)
            tree.insert("", i, text=i-100, values=(nm, fnm, lnm, mob, std,gend)),
        tree.pack()
        print("---------------------------------refreshed------------------------------------")
        messagebox.showinfo("Refreshed", "Succesfully Refreshed")
           
    tree["columns"] = ("one", "two", "three","four","five","six")
    tree.column("one", width=100)
    tree.column("two", width=100)
    tree.column("three", width=100)
    tree.column("four", width=100)
    tree.column("five", width=100)
    tree.column("six", width=100)

    tree.heading("#0", text='ID', anchor='w')
    tree.column("#0", anchor="w")
    tree.heading("one", text="Roll No")
    tree.heading("two", text="Subject 1")
    tree.heading("three", text="Subject 1")
    tree.heading("four", text="Subject 1")
    tree.heading("five", text="Subject 1")
    tree.heading("six",text="Total")
    
        
    Button(feesframe,text="Insert Marks",command=put).grid(row=6,column=0)
    Button(feesframe,text="Remove Marks",command=rem).grid(row=6,column=1)
    Button(feesframe,text="Refresh Marks",command=content).grid(row=6,column=2)
    win2.mainloop()
    
def mainmenu():
    root.withdraw()
    global win 
    win=Toplevel(root)
    win.geometry("1350x270+0+170")
    menu = Menu(win) 
    win.config(menu=menu) 
    filemenu = Menu(menu) 
    menu.add_cascade(label='Options', menu=filemenu) 
    filemenu.add_command(label='Login',command='root.deiconify') 
    filemenu.add_command(label='Register',command='root.deiconify')
    filemenu.add_command(label='Refresh Details',command='content()')
    filemenu.add_separator() 
    filemenu.add_command(label='Exit', command=win.quit) 
    helpmenu = Menu(menu)
    menu.add_cascade(label='Help', menu=helpmenu) 
    helpmenu.add_command(label='About') 
    rightframe=Frame(win)
    win.title('Student Management Systems')
    tree = ttk.Treeview(rightframe)
    tree.pack(padx=5, pady=10)
    win.configure(bg='#E8DAEF')
    frame=Frame(win)
    rightframe.pack(side='left')
    frame.pack(side='left')
    frame.configure(bg='#E8DAEF')
    frame1=Frame(frame) 
    frame1.grid(row='7')
    frame1.configure(bg='#E8DAEF')
    
    rightframe.pack(side='right')
    rightframe.configure(bg='#E8DAEF')
    roll_no=IntVar()
    f_name=StringVar()
    l_name=StringVar()
    mobile_l=IntVar()
    rollentry=Entry(frame,width=8,textvariable=roll_no)
    fnameentry=Entry(frame,width=8,textvariable=f_name)
    lnameentry=Entry(frame,width=8,textvariable=l_name)
    genderentry=ttk.Combobox(frame,values=["Male","Female"],width=7)
    print(dict(genderentry))
    
    
    mobileentry=Entry(frame,width=8,textvariable=mobile_l)
    standardentry=ttk.Combobox(frame,values=["fybsc","sybsc","tybsc"],width=7)
    rollentry.delete(0,END)
    fnameentry.delete(0,END)
    lnameentry.delete(0,END)
    mobileentry.delete(0,END)

    rollentry.grid(row=0,column=1)
    fnameentry.grid(row=1,column=1)
    lnameentry.grid(row=2,column=1)
    genderentry.grid(row=3,column=1)
    mobileentry.grid(row=4,column=1)
    standardentry.grid(row=5,column=1)
    labelfont=('serif',20,'bold')
    l1=Label(frame,text='Roll No:')
    l1.grid(row=0,column=0,sticky=W)
    l2=Label(frame,text='First Name:')
    l2.grid(row=1,column=0,sticky=W)
    l3=Label(frame,text='Last Name:')
    l3.grid(row=2,column=0,sticky=W)
    l6=Label(frame,text='Gender:')
    l6.grid(row=3,column=0,sticky=W)
    
    l4=Label(frame,text='Mobile:')
    l4.grid(row=4,column=0,sticky=W)
    l5=Label(frame,text='Standard:')
    l5.grid(row=5,column=0,sticky=W)
    l1.config(font=labelfont,bg='#E8DAEF')
    l2.config(font=labelfont,bg='#E8DAEF')
    l3.config(font=labelfont,bg='#E8DAEF')
    l4.config(font=labelfont,bg='#E8DAEF')
    l5.config(font=labelfont,bg='#E8DAEF')
    l6.config(font=labelfont,bg='#E8DAEF')
    rollentry.config(font=labelfont)
    fnameentry.config(font=labelfont)
    lnameentry.config(font=labelfont)
    mobileentry.config(font=labelfont)
    standardentry.config(font=labelfont)
    genderentry.config(font=labelfont)


    def put(*args):
        try:
            roll=roll_no.get()
            fname=f_name.get()
            lname=l_name.get()
            mobile=mobile_l.get()
            standard=standardentry.get()
            gen=genderentry.get()
            cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
            cur=cn.cursor()
            cur.execute("insert into student2(roll,first_name,last_name,mobile_no,standard,gender)\
                    values('"+str(roll)+"','"+fname+"','"+lname+"','"+str(mobile)+"','"+standard+"','"+gen+"')")
            cur.execute("commit")
            cn.close()
            rollentry.delete(0,END)
            fnameentry.delete(0,END)
            lnameentry.delete(0,END)
            mobileentry.delete(0,END)
            standardentry.delete(0,END)
            
        except:
            messagebox.showinfo("Error X", "Enter All Valid Feilds!")
    def updt(*args):
        try:
            roll=roll_no.get()
            fname=f_name.get()
            lname=l_name.get()
            mobile=mobile_l.get()
            standard=standardentry.get()
            gender=genderentry.get()
            cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
            cur=cn.cursor()
            cur.execute("UPDATE student2 SET first_name ='"+fname+"',last_name ='"+lname+"', mobile_no = '"+str(mobile)+"',standard ='"+standard+"',gender = '"+(gender)+"' WHERE  roll = '"+str(roll)+"';")
            cur.execute("commit")
            cn.close()
            rollentry.delete(0,END)
            fnameentry.delete(0,END)
            lnameentry.delete(0,END)
            mobileentry.delete(0,END)
            standardentry.delete(0,END)
            
        except:
            messagebox.showinfo("Error X", "cannot update values!\nenter valid details!")
    
    def rem(*args):
        try:
            roll=roll_no.get()
            cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
            cur=cn.cursor()
            cur.execute("DELETE FROM student2 WHERE roll ='"+str(roll)+"'")
            cur.execute("commit")
            cn.close()
            rollentry.delete(0,END)
            fnameentry.delete(0,END)
            lnameentry.delete(0,END)
            mobileentry.delete(0,END)
            standardentry.delete(0,END)
        except:
            messagebox.showinfo("Error X", "Enter The ROll NO !")
    cur.execute("select * from student2")
    rec=cur.fetchall()
    tmp=0
    def content(*args):
        cn=mc.connect(user="nadeem",passwd="Nadeem@2001",host="localhost",database="Student_management")
        cur=cn.cursor()
        cur.execute("select * from student2")
        rec=cur.fetchall()[0]
        p=100
        for row in rec:
            p+=1
        tmp=p+1
    
        for i in tree.get_children():
            tree.delete(i)
        for i in range(101, p+1):
            cur.execute("""select roll from student2 where roll =%s""", (i,))
            nm = cur.fetchone()[0]
            cur.execute("""select first_name from student2 where roll =%s""", (i,))
            fnm = cur.fetchone()[0]
            cur.execute("""select last_name from student2 where roll =%s""",(i,))
            lnm = cur.fetchone()[0]
            cur.execute("""select mobile_no from student2 where roll =%s""", (i,))
            mob = cur.fetchone()[0]
            cur.execute("""select standard from student2 where roll =%s""",(i,))
            std = cur.fetchone()[0]
            cur.execute("""select gender from student2 where roll =%s""",(i,))
            gend=cur.fetchone()[0]
            tree.insert("", i, text=i-100, values=(nm, fnm, lnm, mob, std,gend)),
        tree.pack()
        print("---------------------------------refreshed------------------------------------")
        messagebox.showinfo("Refreshed", "Succesfully Refreshed")

    tree["columns"] = ("one", "two", "three","four","five","six")
    tree.column("one", width=100)
    tree.column("two", width=100)
    tree.column("three", width=100)
    tree.column("four", width=100)
    tree.column("five", width=100)
    tree.column("six", width=100)

    tree.heading("#0", text='ID', anchor='w')
    tree.column("#0", anchor="w")
    tree.heading("one", text="Roll No")
    tree.heading("two", text="Frist Name")
    tree.heading("three", text="Last Name")
    tree.heading("four", text="Mobile")
    tree.heading("five", text="Standard")
    tree.heading("six",text="Gender")


    Button(frame1,text="INSERT",command=put).pack(side='left',pady=2,padx=5)
    Button(frame1,text="REMOVE",command=rem).pack(side='left',pady=2,padx=5)
    Button(frame1,text="UPDATE",command=updt).pack(side='left',pady=2,padx=5)
    Button(frame1,text="refresh",command=content).pack(side='left',pady=2,padx=5)
    Button(frame1,text="result",command=feestk).pack(side='left',pady=2,padx=5)
    Button(frame1,text="export to excel",command=exportxl).pack(side='left',pady=2,padx=5)

    win.mainloop()
    




root=tk.Tk()  
root.geometry("261x70+450+300")     
frame=Frame(root)
frame.pack(fill=X)
root.title('login!')
    
login=StringVar()
password=StringVar()
loginentry=Entry(frame,width=20,textvariable=login)
passentry=Entry(frame,textvariable=password,show="*")
l1=Label(frame,text='UserName:')
l1.grid(row=0,column=0,sticky=W)
l2=Label(frame,text='Password:')
l2.grid(row=1,column=0,sticky=W)
loginentry.grid(row=0,column=1)
passentry.grid(row=1,column=1)

user_pass = {"admin":"admin","nadeem":"nadeem2001"}
def rog():
    user = login.get()
    pwd = password.get()
    file = open("password.txt", "a+")
    file.write("\n" + user + "\n")
    file.write(pwd)
    file.close()
    print("registration completed,please login")


def verify():
    user=login.get()
    pwd=password.get()
    file1 = open("password.txt", "r+")
    verify = file1.read().splitlines()
    if user in verify and pwd in verify:
        print("success")
        mainmenu()

    else:
        messagebox.showinfo("error", "incorrect login details!")

button4=Button(frame,text="Login",width=17,command=verify).grid(row=2,column=1,sticky=W)
button5=Button(frame,text="Register",width=17,command=rog).grid(row=2,column=0,sticky=W,padx=2)
cur=cn.cursor()
root.mainloop()
