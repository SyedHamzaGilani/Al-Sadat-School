from tkinter import *
from tkinter import ttk as tk
import tkinter.messagebox as tmsg
import cx_Oracle
import ctypes


class Pending:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1335x500+15+110')
        self.root.resizable(False, False)
        self.root.title('Pending Dues Details')
        self.root.config(bg = "Cyan")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.frame1 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame1.place(x = 5, y = 10, width = 1325, height = 50)
        Label(self.frame1, text = "Pending Dues Details", font = "Helvatica 20 bold").pack(padx = 5)
        self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame2.place(x = 4, y = 70, width = 920+405, height = 420)
        Label(self.frame2, text = "Search By", font = "Courier 20 bold").place(x = 50, y = 5)
        self.SearchBy = StringVar()
        self.Search = StringVar()
        searchBy = tk.Combobox(self.frame2, textvariable = self.SearchBy, font = "Courier 11 bold")
        searchBy['values'] = ['Student ID', 'Course ID']
        searchBy.place(x = 250, y = 9, width = 305)
        self.SearchE = Entry(self.frame2, textvariable = self.Search, font = "Courier 11 bold")
        self.ValidateSearchBar()
        self.SearchE.place(x = 570, y = 10, width = 305)
        Button(self.frame2, text = "Search", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.SearchCourse).place(x = 590+400, y = 10, height = 25, width = 150)
        Button(self.frame2, text = "Show All", font = "Helvatica 14 bold", bd = 2, relief = GROOVE, bg="#b3b3b3", cursor="hand2", command = self.FetchCourses).place(x = 750+400, y = 10, height = 25, width = 150)
        scrollX = Scrollbar(self.frame2, orient=HORIZONTAL)
        scrollX.place(x=5, y=385, width=880+405)
        scrollY = Scrollbar(self.frame2, orient=VERTICAL)
        scrollY.place(x=890+405, y=50, height=330)
        self.Tree = tk.Treeview(self.frame2, xscrollcommand=scrollX.set, yscrollcommand=scrollY.set, column=('SID', 'CID', 'Fees', 'PenFees', 'DDate'))
        self.Tree.heading('SID', text='Student Id')
        self.Tree.heading('CID', text='Course Id')
        self.Tree.heading('Fees', text='Fees Deposited')
        self.Tree.heading('PenFees', text='Pending Dues')
        self.Tree.heading('DDate', text='Deposit Date')
        self.Tree.column('SID', width=80)
        self.Tree.column('CID', width=80)
        self.Tree.column('Fees', width=80)
        self.Tree.column('PenFees', width=80)
        self.Tree.column('DDate', width=80)
        self.Tree['show'] = ['headings']
        self.Tree.place(x=5, y=50, width=880+405, height=330)
        scrollX.config(command=self.Tree.xview)
        scrollY.config(command=self.Tree.yview)
        self.FetchCourses()

    
    def connection(self):
        connection = cx_Oracle.connect("system/69-Gilani-53@localhost:1521/orcl")
        return connection
    
    
    
    def FetchCourses(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT SID, CID, FEES, PENFEES, TO_CHAR(DDATE, 'MM/DD/YY') FROM ENROLLEDCOURSE WHERE PENFEES > 0 ORDER BY SID ASC"
        cursor.execute(query)
        rows = cursor.fetchall()
        if len(rows) != 0:
            self.Tree.delete(*self.Tree.get_children())
            for row in rows:
                self.Tree.insert("", END, values=row)
            connection.commit()
        cursor.close()
    
        
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

if __name__ == '__main__':
    root = Tk()
    pending = Pending(root)
    root.mainloop()