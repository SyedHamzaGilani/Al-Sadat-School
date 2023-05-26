
from tkinter import *
from tkinter import ttk as tk
import tkinter.messagebox as tmsg
import cx_Oracle
import ctypes
import pandas as pd


class Expense:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1335x500+15+110')
        self.root.resizable(False, False)
        self.root.title('Expense Details')
        self.root.config(bg = "Cyan")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.frame1 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame1.place(x = 5, y = 10, width = 1325, height = 50)
        Label(self.frame1, text = "Expense Dues Details", font = "Helvatica 20 bold").pack(padx = 5)
        self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame2.place(x = 4, y = 70, width = 710, height = 420)
        scrollX = Scrollbar(self.frame2, orient=HORIZONTAL)
        scrollX.place(x=5, y=385, width=655)
        scrollY = Scrollbar(self.frame2, orient=VERTICAL)
        scrollY.place(x=665, y=10, height=330)
        self.Tree = tk.Treeview(self.frame2, xscrollcommand=scrollX.set, yscrollcommand=scrollY.set, column=('date', 'f_rec'))
        self.Tree.heading('date', text='Deposit Date')
        self.Tree.heading('f_rec', text='Fees Received')
        self.Tree.column('date', width = 10)
        self.Tree.column('f_rec', width = 10)
        self.Tree['show'] = ['headings']
        self.Tree.place(x=5, y=10, width=655, height=370)
        scrollX.config(command=self.Tree.xview)
        scrollY.config(command=self.Tree.yview)
        self.frame3 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame3.place(x = 720, y = 70, width = 610, height = 420)
        Label(self.frame3, text = "Total Fees Received = ", font = "Helvatica 20 bold", fg = "Blue").place(x = 50, y = 50)
        self.RecLabel = Label(self.frame3, text = "", font = "Helvatica 20 bold", fg = "Blue")
        self.ReceivedFees()
        self.RecLabel.place(x = 400, y = 50)
        Label(self.frame3, text = "Sum Of Fees Pending = ", font = "Helvatica 20 bold", fg = "Blue").place(x = 50, y = 150)
        self.PenLabel = Label(self.frame3, text = "", font = "Helvatica 20 bold", fg = "Blue")
        self.PendingFees()
        self.PenLabel.place(x = 400, y = 150)
        self.FetchCourses()
        Label(self.frame3, text = "Total Fees = ", font = "Helvatica 20 bold", fg = "Blue").place(x = 50, y = 250)
        total_fees = int(self.PenFees) + int(self.RecFees)
        self.TotLabel = Label(self.frame3, text = f"{total_fees}", font = "Helvatica 20 bold", fg = "Blue")
        self.TotLabel.place(x = 400, y = 250)
        self.FetchCourses()
        Button(self.frame3, text = "Download CSV Files", font = "Helvatica 20 bold", bd = 5, relief = GROOVE, command = self.DownloadCSV).place(x = 300, y = 330)

    
    def connection(self):
        connection = cx_Oracle.connect("system/69-Gilani-53@localhost:1521/orcl")
        return connection
    
    
    
    def FetchCourses(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT TO_CHAR(DDATE, 'DD - MONTH- YYYY'), SUM(FEES) FROM ENROLLEDCOURSE GROUP BY DDATE ORDER BY DDATE ASC"
        cursor.execute(query)
        rows = cursor.fetchall()
        if len(rows) != 0:
            self.Tree.delete(*self.Tree.get_children())
            for row in rows:
                self.Tree.insert("", END, values=row)
            connection.commit()
        cursor.close()
        
    def DownloadCSV(self):
        try:
            query = "SELECT TO_CHAR(DDATE, 'DD - MONTH- YYYY') AS DEPOSITE_DATE, SUM(FEES) AS FEES_RECEIVED FROM ENROLLEDCOURSE GROUP BY DDATE ORDER BY DDATE ASC"
            query1 = "SELECT SID, CID, FEES AS DEPOSITED_FEES, PENFEES AS PENDING_FEES, TO_CHAR(DDATE, 'DD - MONTH- YYYY') AS DEPOSITED_DATE FROM ENROLLEDCOURSE WHERE PENFEES > 0 ORDER BY SID ASC"
            try:
                df = pd.read_sql_query(query, self.connection())
                df1 = pd.read_sql(query1, self.connection())
            except:
                pass
            df.to_excel("Expenses.xlsx", index=False)
            df1.to_excel("Pending.xlsx", index = False)
            tmsg.showinfo("Saved", "Excel Files Are Saved In Your Repository")
        except Exception as e:
            tmsg.showerror("Error", f"Error Dur To {e}")
    
    def PendingFees(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT SUM(PENFEES) AS PENFEES FROM ENROLLEDCOURSE WHERE PENFEES > 0"
        cursor.execute(query)
        rows = cursor.fetchone()
        for i in rows:
            self.PenFees = i
        self.PenLabel.config(text = f"{self.PenFees}")
        connection.commit()
        cursor.close()
    
    def ReceivedFees(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT SUM(FEES) AS FEES FROM ENROLLEDCOURSE"
        cursor.execute(query)
        rows = cursor.fetchone()
        for i in rows:
            self.RecFees = i
        self.RecLabel.config(text = f"{self.RecFees}")
        connection.commit()
        cursor.close()
    
        
    
if __name__ == '__main__':
    root = Tk()
    expense = Expense(root)
    root.mainloop()