from tkinter import *
from tkinter import ttk as tk
from datetime import date
import tkinter.messagebox as tmsg
import tkcalendar
import cx_Oracle
import ctypes


class Attendence:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1335x500+15+110')
        self.root.resizable(False, False)
        self.root.title('Attendence Details')
        self.root.config(bg = "Cyan")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.frame1 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame1.place(x = 5, y = 10, width = 1325, height = 50)
        Label(self.frame1, text = "Add Attendance Details Here", font = "Helvatica 20 bold").pack(padx = 5)
        self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame2.place(x = 5, y = 70, width = 400, height = 360)
        self.CID = StringVar()
        self.SID = StringVar()
        self.Date = StringVar()
        today = date.today()
        today = today.strftime("%m/%d/%y")
        self.Date.set(today)
        self.Status = StringVar()
        self.Status.set("Present")
        Label(self.frame2, text = "Attendance Details", font = "Helvatica 20 bold underline").place(x = 70, y = 10)
        Label(self.frame2, text = "Course ID", font = "Courier 16 bold").place(x = 5, y = 90)
        connection = self.connection()
        cursor = connection.cursor()
        query = 'SELECT CID FROM COURSES ORDER BY CID ASC'
        cursor.execute(query)
        row = cursor.fetchall()
        cidE = tk.Combobox(self.frame2, textvariable = self.CID, font = "Courier 12 bold")
        cidE['values'] = row
        cidE.place(x = 190, y = 90, width = 190)
        Label(self.frame2, text = "Student ID", font = "Courier 16 bold").place(x = 5, y = 140)
        cursor1 = connection.cursor()
        query1 = "SELECT SID FROM STUDENTS ORDER BY SID ASC"
        cursor1.execute(query1)
        row1 = cursor1.fetchall()
        sidE = tk.Combobox(self.frame2, textvariable = self.SID, font = "Courier 12 bold")
        sidE['values'] = row1
        sidE.place(x = 190, y = 140, width = 190)
        Label(self.frame2, text = "Date", font = "Courier 16 bold").place(x = 5, y = 190)
        Entry(self.frame2, textvariable = self.Date, font = "Courier 13 bold").place(x = 190, y = 190, width = 190)
        Button(self.frame2, text="Cal", font="Conicsansms 10 bold", fg="Black", command=self.getDate).place(x=345, y=192,height=20)
        Label(self.frame2, text = "Status", font = "Courier 16 bold").place(x = 5, y = 240)
        self.StatusE = tk.Combobox(self.frame2, textvariable = self.Status, font = "Courier 12 bold")
        self.StatusE['values'] = ["Present", "Absent"]
        self.CheckStatus()
        self.StatusE.place(x = 190, y = 240, width = 190)
        self.frame3 = Frame(self.root, bd = 5, relief = SUNKEN)
        Button(self.frame3, text="Add", font="Helvatica 14 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.InsertData).place(x=5, y=7, height=30)
        Button(self.frame3, text="Update", font="Helvatica 14 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.UpdateData).place(x=65, y=7, height=30)
        Button(self.frame3, text="Delete", font="Helvatica 14 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.DeleteData).place(x=150, y=7, height=30)
        Button(self.frame3, text="Clear", font="Helvatica 16 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command = self.Clear).place(x=230, y=7, height=30, width=80)
        Button(self.frame3, text="Exit", font="Helvatica 16 bold", bd=2, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.root.destroy).place(x=315, y=7, height=30, width=70)
        self.frame3.place(x = 5, y = 440, width = 400, height = 50)
        self.frame4 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame4.place(x = 410, y = 70, width = 920, height = 420)
        Label(self.frame4, text = "Search By", font = "Courier 16 bold").place(x = 5, y = 10)
        self.SearchBy = StringVar()
        self.Search = StringVar()
        searchbyE = tk.Combobox(self.frame4, textvariable = self.SearchBy, font = "Courier 11 bold")
        searchbyE['values'] = ['Course ID', 'Student ID', 'Date']
        searchbyE.place(x = 150, y = 10)
        self.SearchE = Entry(self.frame4, textvariable = self.Search, font = "Courier 11 bold")
        self.ValidateSearchBar()
        self.SearchE.place(x = 370, y = 10)
        Button(self.frame4, text = "Search", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.SearchData).place(x = 590, y = 10, height = 25, width = 150)
        Button(self.frame4, text = "Show All", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.FetchData).place(x = 750, y = 10, height = 25, width = 150)
        scrollX = Scrollbar(self.frame4, orient=HORIZONTAL)
        scrollX.place(x=5, y=385, width=880)
        scrollY = Scrollbar(self.frame4, orient=VERTICAL)
        scrollY.place(x=890, y=50, height=330)
        self.Tree = tk.Treeview(self.frame4, xscrollcommand=scrollX.set, yscrollcommand=scrollY.set, column=('SID', 'CID', 'Date', 'Status'))
        self.Tree.heading('SID', text='Student Id')
        self.Tree.heading('CID', text='Course Id')
        self.Tree.heading('Date', text='Date')
        self.Tree.heading('Status', text='Status')
        self.Tree.column('SID', width=80)
        self.Tree.column('CID', width=80)
        self.Tree.column('Date', width=80)
        self.Tree.column('Status', width=80)
        self.Tree['show'] = ['headings']
        self.Tree.place(x=5, y=50, width=880, height=330)
        scrollX.config(command=self.Tree.xview)
        scrollY.config(command=self.Tree.yview)
        self.Tree.bind('<Double-Button-1>', self.getCursor)
        self.FetchData()

    
    def ExtractEntry(self, event):
        self.Date.set(self.cal.get_date())
        self.cal.destroy()
    
    def getDate(self):
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
        self.root.bind("<Double-Button-1>", self.ExtractEntry)
    
    def CheckStatus(self):
        if self.StatusE.get() == "Present" or self.StatusE.get() == 'Absent':
            # self.StatusE.config(state = "normal")
            pass
        else:
            # self.StatusE.config(state = "disable")
            tmsg.showerror("Status Error", "Student Must Be Present Or Absent")
            self.Status.set("")
        self.StatusE.after(5000, self.CheckStatus)

    def Clear(self):
        self.CID.set("")
        self.SID.set("")
        today = date.today()
        today = today.strftime("%m/%d/%y")
        self.Date.set(today)
        self.Status.set("Present")
        self.Search.set("")
        self.SearchBy.set("")
    
    def connection(self):
        connection = cx_Oracle.connect("system/69-Gilani-53@localhost:1521/orcl")
        return connection
    
    def InsertData(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            sid = self.SID.get()
            cid = self.CID.get()
            adate = self.Date.get()
            status = self.Status.get()
            query = f"INSERT INTO ATTENDENCE (SID, CID, ADATE, STATUS) VALUES ('{sid}', '{cid}', TO_DATE('{adate}', 'MM/DD/YY'), '{status}')"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            tmsg.showinfo("Good Job", f"Attendence Of Student # {self.SID.get()} In Course # {self.CID.get()} Is Inserted Successfully")
            self.FetchData()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Insertion Error", f"Check SID, CID Or Error Due To {e}")
    
    def DeleteData(self):
        a = tmsg.askyesnocancel("Attention", "Do You Want To Delete The Attendence Of Student Permenantly?")
        if a == True:
            try:
                connection = self.connection()
                cursor = connection.cursor()
                query = f"DELETE FROM ATTENDENCE WHERE SID = '{self.SID.get()}' AND CID = '{self.CID.get()}'"
                cursor.execute(query)
                connection.commit()
                cursor.close()
                tmsg.showinfo("Good Job", f"Attendence Of Student # {self.SID.get()} In Course # {self.CID.get()} Is Deleted Successfully")
                self.FetchData()
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
        self.Date.set(row[2])
        self.Status.set(row[3])

    def FetchData(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT SID, CID, TO_CHAR(ADATE, 'MM/DD/YY'), STATUS FROM ATTENDENCE ORDER BY SID ASC"
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
            query = f"UPDATE ATTENDENCE SET ADATE = TO_DATE('{self.Date.get()}', 'MM/DD/YY'), STATUS = '{self.Status.get()}' WHERE SID = '{self.SID.get()}' AND CID = '{self.CID.get()}'"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            tmsg.showinfo("Good Job", f"Attendence Of Student # {self.SID.get()} In Course # {self.CID.get()} Is Updated Successfully")
            self.FetchData()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Updation Error", f"Check SID, CID Or Error Due To {e}")
        
    def SearchData(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            if self.SearchBy.get() == "Course ID":
                query = f"SELECT SID, CID, TO_CHAR(ADATE, 'MM/DD/YY'), STATUS FROM ATTENDENCE WHERE CID = '{self.Search.get()}'"
            elif self.SearchBy.get() == "Student ID":
                query = f"SELECT SID, CID, TO_CHAR(ADATE, 'MM/DD/YY'), STATUS FROM ATTENDENCE WHERE SID = '{self.Search.get()}'"
            elif self.SearchBy.get() == 'Date':
                query = f"SELECT SID, CID, TO_CHAR(ADATE, 'MM/DD/YY'), STATUS FROM ATTENDENCE WHERE ADATE = TO_DATE('{self.Search.get()}', 'MM/DD/YY')"
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
        if self.SearchBy.get() == 'Student ID' or self.SearchBy.get() == "Course ID" or self.SearchBy.get() == "Date":
            self.SearchE.config(state = "normal")
        else:
            self.SearchE.config(state = "disabled")
        self.SearchE.after(500, self.ValidateSearchBar)

    
if __name__ == '__main__':
    root = Tk()
    attendace = Attendence(root)
    root.mainloop()