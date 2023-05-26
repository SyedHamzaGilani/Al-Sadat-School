from tkinter import *
from tkinter import ttk as tk
import tkinter.messagebox as tmsg
import cx_Oracle
import ctypes


class Student:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1335x500+15+110')
        self.root.resizable(False, False)
        self.root.title('Student Details')
        self.root.config(bg = "Cyan")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.frame1 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame1.place(x = 5, y = 10, width = 1325, height = 50)
        Label(self.frame1, text = "Add Student Details Here", font = "Helvatica 20 bold").pack(padx = 5)
        self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame2.place(x = 5, y = 70, width = 400, height = 360)
        self.SID = StringVar()
        self.SName = StringVar()
        self.Grade = StringVar()
        self.FName = StringVar()
        self.PhoneNo = StringVar()
        self.Gmail = StringVar()
        self.Gmail.set("abc@xyzmail.com")
        self.Address = StringVar()

        Label(self.frame2, text = "Student Details", font = "Helvatica 20 bold underline").place(x = 90, y = 10)
        Label(self.frame2, text = "Student ID", font = "Courier 16 bold").place(x = 5, y = 70)
        Entry(self.frame2, textvariable = self.SID, font = "Courier 13 bold").place(x = 190, y = 70, width = 190)
        Label(self.frame2, text = "Student Name", font = "Courier 16 bold").place(x = 5, y = 110)
        Entry(self.frame2, textvariable = self.SName, font = "Courier 13 bold").place(x = 190, y = 110, width = 190)
        Label(self.frame2, text = "Grade", font = "Courier 16 bold").place(x = 5, y = 150)
        Entry(self.frame2, textvariable = self.Grade, font = "Courier 13 bold").place(x = 190, y = 150, width = 190)
        Label(self.frame2, text = "Father Name", font = "Courier 16 bold").place(x = 5, y = 190)
        Entry(self.frame2, textvariable = self.FName, font = "Courier 13 bold").place(x = 190, y = 190, width = 190)
        Label(self.frame2, text = "Phone No", font = "Courier 16 bold").place(x = 5, y = 230)
        Entry(self.frame2, textvariable = self.PhoneNo, font = "Courier 13 bold").place(x = 190, y = 230, width = 190)
        Label(self.frame2, text = "Gmail ID", font = "Courier 16 bold").place(x = 5, y = 270)
        self.GmailE = Entry(self.frame2, textvariable = self.Gmail, font = "Courier 13 bold")
        # self.ValidateGmail()
        self.GmailE.place(x = 190, y = 270, width = 190)
        Label(self.frame2, text = "Address", font = "Courier 16 bold").place(x = 5, y = 310)
        Entry(self.frame2, textvariable = self.Address, font = "Courier 13 bold").place(x = 190, y = 310, width = 190)

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
        searchBy = tk.Combobox(self.frame4, textvariable = self.SearchBy, font = "Courier 11 bold")
        searchBy['values'] = ['SID', 'Student Name', 'Gmail']
        searchBy.place(x = 150, y = 10)
        self.SearchE = Entry(self.frame4, textvariable = self.Search, font = "Courier 11 bold")
        self.ValidateSearchBar()
        self.SearchE.place(x = 370, y = 10)
        Button(self.frame4, text = "Search", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.SearchStudent).place(x = 590, y = 10, height = 25, width = 150)
        Button(self.frame4, text = "Show All", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.FetchStudents).place(x = 750, y = 10, height = 25, width = 150)
        scrollX = Scrollbar(self.frame4, orient=HORIZONTAL)
        scrollX.place(x=5, y=385, width=880)
        scrollY = Scrollbar(self.frame4, orient=VERTICAL)
        scrollY.place(x=890, y=50, height=330)
        self.Tree = tk.Treeview(self.frame4, xscrollcommand=scrollX.set, yscrollcommand=scrollY.set, column=('SID', 'SName', 'Grade', 'FName', 'PhoneNo', 'Gmail', 'Address'))
        self.Tree.heading('SID', text='Student Id')
        self.Tree.heading('SName', text='Student Name')
        self.Tree.heading('Grade', text='Grade')
        self.Tree.heading('FName', text='Father Name')
        self.Tree.heading('PhoneNo', text='Phone No')
        self.Tree.heading('Gmail', text='Gmail ID')
        self.Tree.heading('Address', text='Address')
        self.Tree.column('SID', width=80)
        self.Tree.column('SName', width=80)
        self.Tree.column('Grade', width=80)
        self.Tree.column('FName', width=80)
        self.Tree.column('PhoneNo', width=80)
        self.Tree.column('Gmail', width=80)
        self.Tree.column('Address', width=80)
        self.Tree['show'] = ['headings']
        self.Tree.place(x=5, y=50, width=880, height=330)
        scrollX.config(command=self.Tree.xview)
        scrollY.config(command=self.Tree.yview)
        self.FetchStudents()
        self.Tree.bind("<Double-Button-1>", self.getCursor)


    def Clear(self):
        self.SID.set("")
        self.SName.set("")
        self.Grade.set("")
        self.FName.set("")
        self.PhoneNo.set("")
        self.Gmail.set("abc@xyzmail.com")
        self.Address.set("")
        self.Search.set("")
        self.SearchBy.set("")
    
    # def ValidateGmail(self):
    #     try:
    #         mail = self.GmailE.get()
    #         mail1 = mail.split("@")[0]
    #         mail2 = mail.split("@")[1]
    #     except:
    #         tmsg.showerror("Mail Error", "Please Enter Valid Gmail/Email ID")
    #     self.GmailE.after(5000, self.ValidateGmail)
    
    def connection(self):
        connection = cx_Oracle.connect("system/69-Gilani-53@localhost:1521/orcl")
        return connection
    
    def InsertData(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            sid = self.SID.get()
            name = self.SName.get().upper()
            grade = self.Grade.get().upper()
            fname = self.FName.get().upper()
            pno = self.PhoneNo.get()
            gmail = self.Gmail.get()
            add = self.Address.get().upper()
            query = f"INSERT INTO STUDENTS (SID, SNAME, GRADE, FNAME, PHONENO, GMAIL, ADDRESS) VALUES ('{sid}', '{name}', '{grade}', '{fname}', '{pno}', '{gmail}', '{add}')"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            tmsg.showinfo("Good Job", f"Student # {self.SID.get()} Is Inserted Successfully")
            self.FetchStudents()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Insertion Error", f"Check SID Or Error Due To {e}")
    
    def DeleteData(self):
        a = tmsg.askyesnocancel("Attention", "Do You Want To Delete The Student Details Permenantly?")
        if a == True:
            try:
                connection = self.connection()
                cursor = connection.cursor()
                query = f"DELETE FROM STUDENTS WHERE SID = {self.SID.get()}"
                cursor.execute(query)
                connection.commit()
                cursor.close()
                tmsg.showinfo("Good Job", f"Student # {self.SID.get()} Is Deleted Successfully")
                self.FetchStudents()
                self.Clear()
            except Exception as e:
                tmsg.showerror("Deletion Error", f"Check SID Or Error Due To {e}")
        else:
            pass
    
    def getCursor(self, event):
        cursor = self.Tree.focus()
        item = self.Tree.item(cursor)
        row = item['values']
        self.SID.set(row[0])
        self.SName.set(row[1])
        self.Grade.set(row[2])
        self.FName.set(row[3])
        self.PhoneNo.set(row[4])
        self.Gmail.set(row[5])
        self.Address.set(row[6])

    def FetchStudents(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT * FROM STUDENTS ORDER BY SID ASC"
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
            query = f"UPDATE STUDENTS SET SNAME = '{self.SName.get().upper()}', GRADE = '{self.Grade.get().upper()}', FNAME = '{self.FName.get().upper()}', PHONENO = '{self.PhoneNo.get()}', GMAIL = '{self.Gmail.get()}', ADDRESS = '{self.Address.get().upper()}' WHERE SID = '{self.SID.get()}'"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            tmsg.showinfo("Good Job", f"Student # {self.SID.get()} Is Updated Successfully")
            self.FetchStudents()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Updation Error", f"Check SID Or Error Due To {e}")
        
    def SearchStudent(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            if self.SearchBy.get() == "SID":
                query = f"SELECT * FROM STUDENTS WHERE SID = '{self.Search.get().upper()}'"
            elif self.SearchBy.get() == "Student Name":
                query = f"SELECT * FROM STUDENTS WHERE SNAME = '{self.Search.get().upper()}'"
            else:
                query = f"SELECT * FROM STUDENTS WHERE '{self.SearchBy.get()}' = '{self.Search.get().upper()}'"
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
        if self.SearchBy.get() == 'SID' or self.SearchBy.get() == "Student Name" or self.SearchBy.get() == "Gmail":
            self.SearchE.config(state = "normal")
        else:
            self.SearchE.config(state = "disabled")
        self.SearchE.after(500, self.ValidateSearchBar)

    
if __name__ == '__main__':
    root = Tk()
    Student = Student(root)
    Student.connection()
    root.mainloop()