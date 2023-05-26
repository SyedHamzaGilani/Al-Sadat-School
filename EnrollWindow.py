from tkinter import *
from tkinter import ttk as tk
from datetime import date
import tkcalendar
import tkinter.messagebox as tmsg
import cx_Oracle
from PendingDues import Pending
from ExpensesWindow import Expense
import ctypes


class Enroll:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1335x500+15+110')
        self.root.resizable(False, False)
        self.root.title('Enrollment Details')
        self.root.config(bg = "Cyan")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.menu = Menu(self.root)
        self.submenu = Menu(self.menu, tearoff=0)
        self.submenu.add_command(label="Pending Dues", command = self.PendingWindow)
        self.submenu.add_command(label="Expenses", command = self.ExpenseWindow)
        self.menu.add_cascade(label="File", menu=self.submenu)
        self.root.config(menu = self.menu)
        self.frame1 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame1.place(x = 5, y = 10, width = 1325, height = 50)
        Label(self.frame1, text = "Add Enrollment Details Here", font = "Helvatica 20 bold").pack(padx = 5)
        self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame2.place(x = 5, y = 70, width = 400, height = 360)
        self.SID = StringVar()
        self.CID = StringVar()
        self.Fees = StringVar()
        self.PenFees = StringVar()
        self.DDate = StringVar()
        self.SDate = StringVar()
        self.EDate = StringVar()
        # Label(self.frame2, text = "Enrollment Details", font = "Helvatica 20 bold underline").place(x = 80, y = 10)
        Label(self.frame2, text = "Student ID", font = "Courier 16 bold").place(x = 5, y = 10)
        connection = self.connection()
        self.sid= tk.Combobox(self.frame2, textvariable = self.SID, font = "Courier 12 bold")
        cursor = connection.cursor()
        query = "SELECT SID FROM STUDENTS ORDER BY SID"
        cursor.execute(query)
        rows = cursor.fetchall()
        self.sid['values'] = rows
        self.sid.place(x = 190, y = 10, width = 190)
        Label(self.frame2, text = "Course ID", font = "Courier 16 bold").place(x = 5, y = 60)
        self.cid = tk.Combobox(self.frame2, textvariable = self.CID, font = "Courier 12 bold")
        cursor2 = connection.cursor()
        query = "SELECT CID FROM COURSES ORDER BY CID"
        cursor2.execute(query)
        rows2 = cursor2.fetchall()
        self.cid['values'] = rows2
        self.cid.place(x = 190, y = 60, width = 190)
        Label(self.frame2, text = "Fees", font = "Courier 16 bold").place(x = 5, y = 110)
        Entry(self.frame2, textvariable = self.Fees, font = "Courier 13 bold").place(x = 190, y = 110, width = 190)
        Label(self.frame2, text = "Pending Dues", font = "Courier 16 bold").place(x = 5, y = 155)
        Entry(self.frame2, textvariable = self.PenFees, font = "Courier 13 bold", state = 'disabled').place(x = 190, y = 155, width = 190)
        Button(self.frame2, text="Calculate", font="Conicsansms 10 bold", fg="Black", command=self.CalculatePendings).place(x=305, y=157,height=20)
        Label(self.frame2, text = "Deposit Date", font = "Courier 16 bold").place(x = 5, y = 200)
        Entry(self.frame2, textvariable = self.DDate, font = "Courier 13 bold").place(x = 190, y = 200, width = 190)
        Button(self.frame2, text="Cal", font="Conicsansms 10 bold", fg="Black", command=self.getDDate).place(x=345, y=202,height=20)
        Label(self.frame2, text = "Start Date", font = "Courier 16 bold").place(x = 5, y = 250)
        Entry(self.frame2, textvariable = self.SDate, font = "Courier 13 bold").place(x = 190, y = 250, width = 190)
        Button(self.frame2, text="Cal", font="Conicsansms 10 bold", fg="Black", command=self.getSDate).place(x=345, y=252,height=20)
        Label(self.frame2, text = "Exp End Date", font = "Courier 16 bold").place(x = 5, y = 300)
        Entry(self.frame2, textvariable = self.EDate, font = "Courier 13 bold").place(x = 190, y = 300, width = 190)
        Button(self.frame2, text="Cal", font="Conicsansms 10 bold", fg="Black", command=self.getEDate).place(x=345, y=302,height=20)
        self.frame3 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.Abutton = Button(self.frame3, text="Add", font="Helvatica 14 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.InsertData)
        self.Abutton.place(x=5, y=7, height=30)
        self.Ubutton = Button(self.frame3, text="Update", font="Helvatica 14 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.UpdateData)
        self.Ubutton.place(x=65, y=7, height=30)
        self.Dbutton = Button(self.frame3, text="Delete", font="Helvatica 14 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.DeleteData)
        self.Dbutton.place(x=150, y=7, height=30)
        # self.ValidateEntry()
        Button(self.frame3, text="Clear", font="Helvatica 16 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.Clear).place(x=230, y=7, height=30, width=80)
        Button(self.frame3, text="Exit", font="Helvatica 16 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.root.destroy).place(x=315, y=7, height=30, width=70)
        self.frame3.place(x = 5, y = 440, width = 400, height = 50)
        self.frame4 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame4.place(x = 410, y = 70, width = 920, height = 420)
        Label(self.frame4, text = "Search By", font = "Courier 16 bold").place(x = 5, y = 10)
        self.SearchBy = StringVar()
        self.Search = StringVar()
        searchBy = tk.Combobox(self.frame4, textvariable = self.SearchBy, font = "Courier 11 bold")
        searchBy['values'] = ['Student ID', 'Course ID']
        searchBy.place(x = 150, y = 10)
        self.SearchE = Entry(self.frame4, textvariable = self.Search, font = "Courier 11 bold")
        self.ValidateSearchBar()
        self.SearchE.place(x = 370, y = 10)
        Button(self.frame4, text = "Search", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.SearchCourse).place(x = 590, y = 10, height = 25, width = 150)
        Button(self.frame4, text = "Show All", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.FetchCourses).place(x = 750, y = 10, height = 25, width = 150)
        scrollX = Scrollbar(self.frame4, orient=HORIZONTAL)
        scrollX.place(x=5, y=385, width=880)
        scrollY = Scrollbar(self.frame4, orient=VERTICAL)
        scrollY.place(x=890, y=50, height=330)
        self.Tree = tk.Treeview(self.frame4, xscrollcommand=scrollX.set, yscrollcommand=scrollY.set, column=('SID', 'CID', 'Fees', 'PenFees', 'DDate', 'SDate', 'EDate'))
        self.Tree.heading('SID', text='Student Id')
        self.Tree.heading('CID', text='Course Id')
        self.Tree.heading('Fees', text='Fees')
        self.Tree.heading('PenFees', text='Pending Dues')
        self.Tree.heading('DDate', text='Deposit Date')
        self.Tree.heading('SDate', text='Start Date')
        self.Tree.heading('EDate', text='Exp End Date')
        self.Tree.column('SID', width=80)
        self.Tree.column('CID', width=80)
        self.Tree.column('Fees', width=80)
        self.Tree.column('PenFees', width=80)
        self.Tree.column('DDate', width=80)
        self.Tree.column('SDate', width=80)
        self.Tree.column('EDate', width=80)
        self.Tree['show'] = ['headings']
        self.Tree.place(x=5, y=50, width=880, height=330)
        scrollX.config(command=self.Tree.xview)
        scrollY.config(command=self.Tree.yview)
        self.Tree.bind('<Double-Button-1>', self.getCursor)
        self.FetchCourses()
        # self.CalculatePendings()

    def ExtractSEntry(self, event):
        self.SDate.set(self.cal.get_date())
        self.cal.destroy()
    
    def ExtractEEntry(self, event):
        self.EDate.set(self.cal1.get_date())
        self.cal1.destroy()
    
    def ExtractDEntry(self, event):
        self.DDate.set(self.cal2.get_date())
        self.cal2.destroy()

    def getDDate(self):
        try:
            self.cal2.destroy()
        except:
            pass
        today = date.today()
        year = int(today.strftime('%Y'))
        month = int(today.strftime('%m'))
        day = int(today.strftime('%d'))
        self.cal2 = tkcalendar.Calendar(self.frame2, selectmode='day', year=year, month=month, day=day)
        self.cal2.place(x=130, y=130)
        self.root.bind("<Double-Button-1>", self.ExtractDEntry)

    def getSDate(self):
        try:
            self.cal.destroy()
        except:
            pass
        today = date.today()
        year = int(today.strftime('%Y'))
        month = int(today.strftime('%m'))
        day = int(today.strftime('%d'))
        self.cal = tkcalendar.Calendar(self.frame2, selectmode='day', year=year, month=month, day=day)
        self.cal.place(x=130, y=150)
        self.root.bind("<Double-Button-1>", self.ExtractSEntry)

    def getEDate(self):
        try:
            self.cal1.destroy()
        except:
            pass
        today = date.today()
        year = int(today.strftime('%Y'))
        month = int(today.strftime('%m'))
        day = int(today.strftime('%d'))
        self.cal1 = tkcalendar.Calendar(self.frame2, selectmode='day', year=year, month=month, day=day)
        self.cal1.place(x=130, y=160)
        self.root.bind("<Double-Button-1>", self.ExtractEEntry)
    
    def Clear(self):
        self.SID.set("")
        self.CID.set("")
        self.Fees.set("")
        self.PenFees.set("")
        self.DDate.set("")
        self.SDate.set("")
        self.EDate.set("")
        self.Search.set("")
        self.SearchBy.set("")
    
    def connection(self):
        connection = cx_Oracle.connect("system/69-Gilani-53@localhost:1521/orcl")
        return connection
    
    def CalculatePendings(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            query = f"SELECT CFEES FROM COURSES WHERE CID = '{self.CID.get()}'"
            cursor.execute(query)
            rows = cursor.fetchone()
            connection.commit()
            cursor.close()
            dep = int(self.Fees.get())
            for i in rows:
                fees = int(i)
                break
            pending = fees - dep
            if pending > 0:
                self.PenFees.set(pending)
                self.Abutton.config(state = 'normal')
                self.Ubutton.config(state = 'normal')
                self.Dbutton.config(state = 'normal')
            elif pending < 0:
                tmsg.showwarning("Auto Pending Error", f"You Have Written More Than The Expected Fees\n The Expected Fees is {fees}")
                self.Abutton.config(state = 'disabled')
                self.Ubutton.config(state = 'disabled')
                self.Dbutton.config(state = 'disabled')
            elif pending == 0:
                self.PenFees.set('0')
                self.Abutton.config(state = 'normal')
                self.Ubutton.config(state = 'normal')
                self.Dbutton.config(state = 'normal')
        except Exception as e:
            tmsg.showerror("Auto Pending Error", f"Check Fees Entry, It's Must Be Of Int Data Type Or Error Due To {e}")

    
    def InsertData(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            sid = self.SID.get()
            cid = self.CID.get()
            fees = self.Fees.get()
            pending = self.PenFees.get()
            ddate = self.DDate.get()
            sdate = self.SDate.get()
            edate = self.EDate.get()
            query = f"INSERT INTO ENROLLEDCOURSE (SID, CID, FEES,  PENFEES, DDATE, START_DATE, EXP_END_DATE) VALUES ('{sid}', '{cid}', '{fees}', '{pending}', TO_DATE('{ddate}', 'MM/DD/YY'), TO_DATE('{sdate}', 'MM/DD/YY'), TO_DATE('{edate}', 'MM/DD/YY'))"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            # self.CalculatePendings()
            tmsg.showinfo("Good Job", f"Student # {self.SID.get()} Is Enrolled In Course # {self.CID.get()} Is Inserted Successfully")
            self.FetchCourses()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Insertion Error", f"Check SID, Course ID Or Error Due To {e}")
    
    def DeleteData(self):
        a = tmsg.askyesnocancel("Attention", "Do You Want To Delete The Enrollement Details Permenantly?")
        if a == True:
            try:
                connection = self.connection()
                cursor = connection.cursor()
                query = f"DELETE FROM ENROLLEDCOURSE WHERE SID = '{self.SID.get()}' AND CID = '{self.CID.get()}'"
                cursor.execute(query)
                connection.commit()
                cursor.close()
                tmsg.showinfo("Good Job", f"Student # {self.SID.get()} Who Is Enrolled In Course # {self.CID.get()} Is Deleted Successfully")
                self.FetchCourses()
                self.Clear()
            except Exception as e:
                tmsg.showerror("Deletion Error", f"Check SID, CID Or Error Due To {e}")
        else:
            pass
    
    def getCursor(self, event):
        cursor = self.Tree.focus()
        item = self.Tree.item(cursor)
        row = item['values']
        self.SID.set(row[0])
        self.CID.set(row[1])
        self.Fees.set(row[2])
        self.PenFees.set(row[3])
        self.DDate.set(row[4])
        self.SDate.set(row[5])
        self.EDate.set(row[6])

    def FetchCourses(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT SID, CID, FEES, PENFEES, TO_CHAR(DDATE, 'MM/DD/YY'), TO_CHAR(START_DATE, 'MM/DD/YY'), TO_CHAR(EXP_END_DATE, 'MM/DD/YY') FROM ENROLLEDCOURSE ORDER BY SID ASC"
        cursor.execute(query)
        rows = cursor.fetchall()
        if len(rows) != 0:
            self.Tree.delete(*self.Tree.get_children())
            for row in rows:
                self.Tree.insert("", END, values=row)
            connection.commit()
        cursor.close()
    
    def UpdateData(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            query = f"UPDATE ENROLLEDCOURSE SET FEES = '{self.Fees.get()}', PENFEES = '{self.PenFees.get()}', DDATE = TO_DATE('{self.DDate.get()}', 'MM/DD/YY'), START_DATE = TO_DATE('{self.SDate.get()}', 'MM/DD/YY'), EXP_END_DATE = TO_DATE('{self.EDate.get()}', 'MM/DD/YY') WHERE SID = '{self.SID.get()}' AND CID = '{self.CID.get()}'"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            # self.CalculatePendings()
            tmsg.showinfo("Good Job", f"Student # {self.SID.get()} With Course # {self.CID.get()} Is Updated Successfully")
            self.FetchCourses()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Updation Error", f"Check SID, CID Or Error Due To {e}")
        
    def SearchCourse(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            if self.SearchBy.get() == "Student ID":
                query = f"SELECT SID, CID, FEES, PENFEES, TO_CHAR(DDATE, 'MM/DD/YY'), TO_CHAR(START_DATE, 'MM/DD/YY'), TO_CHAR(EXP_END_DATE, 'MM/DD/YY') FROM ENROLLEDCOURSE WHERE SID = '{self.Search.get()}'"
            elif self.SearchBy.get() == "Course ID":
                query = f"SELECT SID, CID, FEES, PENFEES, TO_CHAR(DDATE, 'MM/DD/YY'), TO_CHAR(START_DATE, 'MM/DD/YY'), TO_CHAR(EXP_END_DATE, 'MM/DD/YY') FROM ENROLLEDCOURSE WHERE CID = '{self.Search.get()}'"
            else:
                query = f"SELECT SID, CID, FEES, PENFEES, TO_CHAR(DDATE, 'MM/DD/YY'), TO_CHAR(START_DATE, 'MM/DD/YY'), TO_CHAR(EXP_END_DATE, 'MM/DD/YY') FROM ENROLLEDCOURSE WHERE '{self.SearchBy.get()}' = '{self.Search.get()}'"
            cursor.execute(query)
            rows = cursor.fetchall()
            if len(rows) != 0:
                self.Tree.delete(*self.Tree.get_children())
                for row in rows:
                    self.Tree.insert("", END, values=row)
            connection.commit()
            cursor.close()
        except Exception as e:
            tmsg.showerror("Searching Error", f"Error Due To {e}")
        
    def ValidateSearchBar(self):
        if self.SearchBy.get() == 'Student ID' or self.SearchBy.get() == "Course ID":
            self.SearchE.config(state = "normal")
        else:
            self.SearchE.config(state = "disabled")
        self.SearchE.after(500, self.ValidateSearchBar)
    
    def ValidateEntry(self):
        connection = self.connection()
        connection2 = self.connection()
        cursor = connection.cursor()
        query = "SELECT SID FROM STUDENTS ORDER BY SID"
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor2 = connection2.cursor()
        query = "SELECT CID FROM COURSES ORDER BY CID"
        cursor2.execute(query)
        rows2 = cursor2.fetchall()
        for i in rows:
            if self.sid.get() == f"{i[0]}":
                self.Abutton.config(state = "normal")
                self.Ubutton.config(state = "normal")
                self.Dbutton.config(state = "normal")
            else:
                self.Abutton.config(state = "disabled")
                self.Dbutton.config(state = "disabled")
                self.Ubutton.config(state = "disabled")
        connection.commit()
        cursor.close()
        for i in rows2:
            if self.cid.get() == f"{i[0]}":
                self.Abutton.config(state = "normal")
                self.Ubutton.config(state = "normal")
                self.Dbutton.config(state = "normal")
            else:
                self.Abutton.config(state = "disabled")
                self.Dbutton.config(state = "disabled")
                self.Ubutton.config(state = "disabled")
        cursor2.close()
        connection2.commit()
        self.sid.after(500, self.ValidateEntry)
    

    def PendingWindow(self):
        try:
            self.newwindow.destroy()
        except:
            pass
        self.newwindow = Toplevel(self.root)
        self.pending = Pending(self.newwindow)
        try:
            self.newwindow1.destroy()
        except:
            pass
    
    def ExpenseWindow(self):
        try:
            self.newwindow1.destroy()
        except:
            pass
        self.newwindow1 = Toplevel(self.root)
        self.expense = Expense(self.newwindow1)
        try:
            self.newwindow.destroy()
        except:
            pass


if __name__ == '__main__':
    root = Tk()
    Enroll = Enroll(root)
    root.mainloop()