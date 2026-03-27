from tkinter import*
import tkinter as ttk
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
import pyodbc






try:
    conn_str=(r'DRIVER={Microsoft Access Driver (*.mdb};'
              r'DBQ=D:\Programming Languages\python\Python project\TDAY\treeview\treeviewhw\treeview1.mdb;')

    conn=pyodbc.connect(conn_str)
    print("connection satablished")
    
    cursor=conn.cursor()    #cursor is a sql command running option
except pyodbc.Error as e:
    print("Data Base Error",e)




def printdata():
    if entindxno.get()=='' or entname.get()=='' or entage.get()=='' :
        tkinter.messagebox.showerror("Data Error","Fill All Blanks")
    else:
        
        no=entindxno.get()
        name=entname.get()
        age=entage.get()
        cursor.execute("insert into Treeview(StNo,StName,StAge)values(?,?,?)",no,name,age)
        tkinter.messagebox.showinfo("Add new data","Add new record Successfully")

        conn.commit()
    
'''def search():
    btnprint["state"]=DISABLED
    no=entindxno.get()
    q="select*from Treeview where stno=?"
    cursor.execute(q,(no,))
    row=cursor.fetchone()
    stno=row[0]
    stname=row[1]
    stage=row[2]
    entname.delete(0,END)    
    entage.delete(0,END)
    entname.insert(END,stname)
    entage.insert(END,stage)

    #print(row)'''

def search():
    btnprint["state"]=DISABLED
    no=entindxno.get()
    try:
        if no != "":
            q="select*from Treeview where stno=?"
            cursor.execute(q,(no,))
            row=cursor.fetchone()
            stno=row[0]
            stname=row[1]
            stage=row[2]
            entname.delete(0,END)    
            entage.delete(0,END)
            entname.insert(END,stname)
            entage.insert(END,stage)
        else:
            tkinter.messagebox.showerror("Error Input","Please Input \n student Number")
            btnprint["state"]=NORMAL
            entindxno.delete(0,END)
            entname.delete(0,END)
            entage.delete(0,END)            

    except:
        tkinter.messagebox.showinfo("Search Record","Input correct \n student Number")
        btnprint["state"]=NORMAL
        entindxno.delete(0,END)
        entname.delete(0,END)
        entage.delete(0,END)        
    
def clear():
    btnprint["state"]=NORMAL
    entindxno.delete(0,END)
    entname.delete(0,END)
    entage.delete(0,END)
    
def exitbtn():
    exitme=tkinter.messagebox.askyesno("Exit","Are you Want To Exit")
    if exitme==1:
        conn.close()
        win.destroy()
    

def update():

    no=entindxno.get()
    name=entname.get()
    age=entage.get()
    cursor.execute("update Treeview set StName=?,StAge=? where StNo=?",(name, age, no))
    tkinter.messagebox.showinfo("Update data","Update Successfully")
    
def display():
    cursor.execute("select * from Treeview")
    result=cursor.fetchall()
    if len(result)!=0:
        memrecords.delete(*memrecords.get_children())
        for row in result[0:]:
            memrecords.insert("",END,values=(row[0],row[1],row[2]))
        
    conn.commit()
        


win=Tk()
win.title("Dpythontreeview")
win.geometry("620x600+450+40")
win.configure(bg="#708090")
win.resizable(0,0)

lblindxno=Label(win,text="Student Number:",font="none 15",bg="#708090",fg="white")
lblindxno.place(x=20,y=50)
entindxno=Entry(win,font="none 15",width=15)
entindxno.place(x=200,y=50)

lblname=Label(win,text="Student Name:",font="none 15",bg="#708090",fg="white")
lblname.place(x=20,y=100)
entname=Entry(win,font="none 15",width=35)
entname.place(x=200,y=100)

lblage=Label(win,text="Student Age:",font="none 15",bg="#708090",fg="white")
lblage.place(x=20,y=150)
entage=Entry(win,font="none 15",width=35)
entage.place(x=200,y=150)

btnprint=Button(win,text="Insert",font="none 15",width="10",bg="darkgray",command=printdata)
btnprint.place(x=100,y=200)

btnsrh=Button(win,text="Search",font="none 15",width="10",bg="darkgray",command=search)
btnsrh.place(x=240,y=200)


btndis=Button(win,text="Display",font="none 15",width="10",height=1,bg="darkgray",command=display)
btndis.place(x=380,y=200)


btnup=Button(win,text="Update",font="none 15",width="10",bg="darkgray",command=update)
btnup.place(x=100,y=250)


btnclr=Button(win,text="Clear",font="none 15",width="10",bg="darkgray",command=clear)
btnclr.place(x=240,y=250)


btnext=Button(win,text="Exit",font="none 15",width="10",bg="darkgray",command=exitbtn)
btnext.place(x=380,y=250)


viewframe=Frame(win,bd=5,width=520,height=150,relief=RIDGE,bg="cadet blue")
viewframe.place(x=10,y=300)

scroll_y=Scrollbar(viewframe,orient=VERTICAL)
memrecords=ttk.Treeview(viewframe,height=10,columns=("stno","stname","stage"),yscrollcommand=scroll_y.set)
scroll_y.pack(side=RIGHT,fill=Y)

memrecords.heading("stno",text="StNo")
memrecords.heading("stname",text="StName",anchor=W)
memrecords.heading("stage",text="StAge")

memrecords["show"]="headings"
memrecords.column("stno",width=90,anchor=CENTER)
memrecords.column("stname",width=380,anchor=W)
memrecords.column("stage",width=100,anchor=CENTER)

memrecords.pack(fill=BOTH,expand=1)


win.mainloop()
