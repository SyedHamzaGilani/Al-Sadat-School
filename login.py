import tkinter.ttk
from tkinter import *
from PIL import Image, ImageTk
import cx_Oracle
import tkinter.messagebox as tmsg
from Account import Account
from main import MainWindow
from threading import Thread
import pyttsx3
import ctypes

class Login:
    def __init__(self, root):
        self.root = root
        self.root.geometry("500x500+400+100")
        self.root.resizable(False, False)
        self.root.title("Login")
        self.c = 0
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        # img1 = Image.open("Images/left.jpg")
        # img1 = img1.resize((500, 500), Image.Resampling.LANCZOS)
        # self.img1 = ImageTk.PhotoImage(img1)
        # img1 = Label(self.root, image=self.img1)
        # img1.place(x=0, y=0)
        self.root.config(bg="Black")
        frame1 = Frame(self.root, bg="#f0f0f0")
        frame1.place(x=5, y=5, width=490, height=490)
        self.label = Label(frame1, text = "Enter Login Details", bg="#f0f0f0", font="Courier 20 bold underline")
        # self.label.place(x = 110, y = 20)
        self.label.pack(side=TOP, pady=20, padx = 30)
        img2 = Image.open("images/user.jpg")
        img2 = img2.resize((30, 30), Image.Resampling.LANCZOS)
        self.img2 = ImageTk.PhotoImage(img2)
        img2 = Label(frame1, image=self.img2)
        img2.place(x=10+50, y=95+40)
        Label(frame1, text="User Name", bg="#f0f0f0", fg="Black", font="Helvatica 12 bold").place(x=50+50, y=100+40)
        self.UserName = StringVar()
        self.PassWord = StringVar()
        self.Check = IntVar()
        Entry(frame1, textvariable=self.UserName, font="Helvatica 12 bold", fg='Black').place(x=170+50, y=100+40)
        img3 = Image.open("images/password.jpg")
        img3 = img3.resize((30, 30), Image.Resampling.LANCZOS)
        self.img3 = ImageTk.PhotoImage(img3)
        img3 = Label(frame1, image=self.img3)
        img3.place(x=10+50, y=165+40)
        Label(frame1, text="Password", bg="#f0f0f0", fg="Black", font="Helvatica 12 bold").place(x=50+50, y=170+40)
        self.PassWordE = Entry(frame1, textvariable=self.PassWord, font="Helvatica 12 bold", fg='Black', show = '*')
        self.PassWordE.place(x=170+50, y=170+40)
        self.CheckE = tkinter.ttk.Checkbutton(frame1, variable = self.Check, onvalue = 1, offvalue = 0, command = self.ShowPassword).place(x = 250, y = 242)
        Label(frame1, text = "Show Password", bg = "#f0f0f0", fg = "Black", font = "Helvatica 12 bold").place(x = 275, y = 240)
        Button(frame1, text="Reset Password", font="Arial 13 bold", bg="#b3b3b3", fg="Black", cursor = "hand2", command=self.Reset).place(x=10, y=240+70, width=385+80)
        Button(frame1, text="Clear Now", font="Arial 13 bold", bg="#b3b3b3", fg="Black", cursor = "hand2", command=self.Clear).place(x=10, y=280+70, width=385+80)
        # Button(frame1, text = "Create New Account", font = "Arial 13 bold", bg = "Black", fg = "#f0f0f0", command = self.newAccount).place(x = 5, y = 300, width = 190)
        # Button(frame1, text = "Login Now", font = "Arial 13 bold", bg = "Black", fg = "#f0f0f0", command = self.Logged).place(x = 200, y = 300, width = 190)
        Button(frame1, text="Login Now", font="Arial 13 bold", bg="#b3b3b3", fg="Black", cursor = "hand2", command=self.Logged).place(x=10, y=320+70, width=385+80)
        Button(frame1, text="Exit", font="Arial 13 bold", bg="#b3b3b3", fg="Black", cursor = "hand2", command=self.loggedOut).place(x=10, y=360+70, width=385+80)
        self.CreatingDatabase()

    def ShowPassword(self):
        if self.Check.get() == 1:
            self.PassWordE.config(show = '')
        else:
            self.PassWordE.config(show = '*')

    def Clear(self):
        self.UserName.set("")
        self.PassWord.set("")
        self.Check.set(0)
        self.PassWordE.config(show = "*")

    def loggedOut(self):
        self.root.destroy()

    def Logged(self):
        try:
            a = str(self.UserName.get().upper())
            b = str(self.PassWord.get())
            connection = self.getConnection()
            cursor = connection.cursor()
            query = "SELECT * FROM LOGIN"
            cursor.execute(query)
            row = cursor.fetchall()
            for rows, rows1 in row:
                if a == rows and b == (rows1):
                    self.Clear()
                    # self.threading()
                    self.root.state("iconic")
                    self.mainWindow()
                    # self.root.state('withdrawn')
                else:
                    tmsg.showerror("Error", "Invalid UserName Or Password")

            connection.commit()
            cursor.close()
        except Exception as es:
            tmsg.showerror("Error", f"Error due to {es}")

    def Reset(self):
        self.newwindow = Toplevel(self.root)
        self.account = Account(self.newwindow)


    def threading(self):
        self.x = Thread(target=self.talking)
        # self.x.start()

    def talking(self):
        engine = pyttsx3.init()
        text = "You Have Successfully Logged In"
        voices = engine.getProperty('voices')
        engine.setProperty('voice', voices[1].id)
        engine.setProperty('rate', 200)
        engine.say(text)
        engine.runAndWait()

    def mainWindow(self):
        self.newwindow1 = Toplevel(self.root)
        self.threading()
        self.main = MainWindow(self.newwindow1)
        self.x.start()

    def getConnection(self):
        connection = cx_Oracle.connect('system/69-Gilani-53@localhost:1521/orcl')
        return connection

    def CreatingDatabase(self):
        try:
            connection1 = self.getConnection()
            cursor1 = connection1.cursor()
            query1 = """
            CREATE TABLE COURSES(
                CID VARCHAR2(50),
                CNAME VARCHAR2(50),
                CFEES VARCHAR2(50),
                CONSTRAINTS COURSES_PK PRIMARY KEY (CID)
            )
            """
            cursor1.execute(query1)
            connection1.commit()
            cursor1.close()
            tmsg.showinfo("Congrats", "Course Database Created Successfully")
        except Exception as e:
            pass
        try:
            connection2 = self.getConnection()
            cursor2 = connection2.cursor()
            query2 = """
            CREATE TABLE STUDENTS(
                SID VARCHAR2(50),
                SNAME VARCHAR2(50),
                GRADE VARCHAR2(10),
                FNAME VARCHAR2(50),
                PHONENO VARCHAR2(13),
                GMAIL VARCHAR2(30),
                ADDRESS VARCHAR2(100),
                CONSTRAINTS STUDENTS_PK PRIMARY KEY (SID)
            )
            """
            cursor2.execute(query2)
            connection2.commit()
            cursor2.close()
            tmsg.showinfo("Congrats", "Student Database Created Successfully")
        except Exception as e:
            pass
        try:
            connection3 = self.getConnection()
            cursor3 = connection3.cursor()
            query3 = """
            CREATE TABLE ENROLLEDCOURSE(
                SID VARCHAR2(50),
                CID VARCHAR2(50),
                FEES VARCHAR2(50),
                PENFEES VARCHAR2(50),
                DDATE DATE,
                START_DATE DATE,
                EXP_END_DATE DATE,
                CONSTRAINTS ENROLLEDCOURSE_PK PRIMARY KEY (SID, CID),
                CONSTRAINTS ENROLLEDCOURSE_FK FOREIGN KEY (SID) REFERENCES STUDENTS (SID),
                CONSTRAINTS ENROLLEDCOURSE_FK1 FOREIGN KEY (CID) REFERENCES COURSES (CID)
            )
            """
            cursor3.execute(query3)
            connection3.commit()
            cursor3.close()
            tmsg.showinfo("Congrats", "Enrollment Database Created Successfully")
        except Exception as e:
            pass
        try:
            connection4 = self.getConnection()
            cursor4 = connection4.cursor()
            query4 = """
            CREATE TABLE LOGIN(
            USERNAME VARCHAR2(100),
            PASSWORD VARCHAR2(100),
            CONSTRAINTS LOGIN_PK PRIMARY KEY (USERNAME)
            )
            """
            cursor4.execute(query4)
            connection4.commit()
            cursor4.close()
            tmsg.showinfo("Congrats", "Login Database Created Successfully")
        except Exception as e:
            pass
        try:
            connection5 = self.getConnection()
            cursor5 = connection5.cursor()
            query5 = "INSERT INTO LOGIN (USERNAME, PASSWORD) VALUES ('SYSTEM', '@System')"
            cursor5.execute(query5)
            connection5.commit()
            cursor5.close()
            tmsg.showinfo("Congrats", "Password Created Successfully")
        except:
            pass
        try:
            connection6 = self.getConnection()
            cursor6 = connection6.cursor()
            query6 = """
            CREATE TABLE ATTENDENCE(
                SID VARCHAR2(50),
                CID VARCHAR2(50),
                ADATE DATE,
                STATUS VARCHAR2(10),
                CONSTRAINTS ATTENDENCE_PK PRIMARY KEY (SID, CID, ADATE),
                CONSTRAINTS ATTENDENCE_FK FOREIGN KEY (SID) REFERENCES STUDENTS (SID),
                CONSTRAINTS ATTENDENCE_FK1 FOREIGN KEY (CID) REFERENCES COURSES (CID)
            )
            """
            cursor6.execute(query6)
            connection6.commit()
            cursor6.close()
            tmsg.showinfo("Congrats", "Attendence Database Created Successfully")
        except Exception as e:
            pass

        try:
            connection7 = self.getConnection()
            cursor7 = connection7.cursor()
            query7 = """
            CREATE TABLE CERTIFICATE(
                CERT_NO VARCHAR2(50),
                CID VARCHAR2(50) NOT NULL,
                SID VARCHAR2(50) NOT NULL,
                CDATE DATE,
                CONSTRAINTS CERTIFICATE_UQ UNIQUE (CID, SID),
                CONSTRAINTS CERTIFICATE_FK FOREIGN KEY (SID) REFERENCES STUDENTS (SID),
                CONSTRAINTS CERTIFICATE_FK1 FOREIGN KEY (CID) REFERENCES COURSES (CID)
            )
            """
            cursor7.execute(query7)
            connection7.commit()
            cursor7.close()
            tmsg.showinfo("Congrats", "Certificate Database Created Successfully")
        except:
            pass


if __name__ == '__main__':
    root = Tk()
    login = Login(root)
    root.mainloop()