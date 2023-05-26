from tkinter import *
from tkinter import ttk as tk
import cx_Oracle
import tkinter.messagebox as tmsg
import ctypes


class Course:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1335x500+15+110')
        self.root.resizable(False, False)
        self.root.title('Course Details')
        self.root.config(bg = "Cyan")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.frame1 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame1.place(x = 5, y = 10, width = 1325, height = 50)
        Label(self.frame1, text = "Add Course Details Here", font = "Helvatica 20 bold").pack(padx = 5)
        self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame2.place(x = 5, y = 70, width = 400, height = 360)
        Label(self.frame2, text = "Course Details", font = "Helvatica 20 bold underline").place(x = 100, y = 10)
        self.courseId = StringVar()
        self.courseName = StringVar()
        self.courseFees = StringVar()
        Label(self.frame2, text = "Course ID", font = "Courier 16 bold").place(x = 5, y = 100)
        Entry(self.frame2, textvariable = self.courseId, font = "Courier 13 bold").place(x = 190, y = 100, width = 190)
        Label(self.frame2, text = "Course Name", font = "Courier 16 bold").place(x = 5, y = 190)
        Entry(self.frame2, textvariable = self.courseName, font = "Courier 13 bold").place(x = 190, y = 190, width = 190)
        Label(self.frame2, text = "Course Fees", font = "Courier 16 bold").place(x = 5, y = 280)
        Entry(self.frame2, textvariable = self.courseFees, font = "Courier 13 bold").place(x = 190, y = 280, width = 190)
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
        searchBy['values'] = ['Course ID', 'Course Name']
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
        self.Tree = tk.Treeview(self.frame4, xscrollcommand=scrollX.set, yscrollcommand=scrollY.set, column=('CourseId', 'CourseName', 'CourseFees'))
        self.Tree.heading('CourseId', text='Course Id')
        self.Tree.heading('CourseName', text='Course Name')
        self.Tree.heading('CourseFees', text='Course Fees')
        self.Tree.column('CourseId', width=80)
        self.Tree.column('CourseName', width=80)
        self.Tree.column('CourseFees', width=80)
        self.Tree['show'] = ['headings']
        self.Tree.place(x=5, y=50, width=880, height=330)
        scrollX.config(command=self.Tree.xview)
        scrollY.config(command=self.Tree.yview)
        self.Tree.bind("<Double-Button-1>", self.getCursor)
        self.FetchCourses()
    
    def Clear(self):
        self.courseId.set("")
        self.courseName.set("")
        self.courseFees.set("")
        self.Search.set("")
        self.SearchBy.set("")
    
    def connection(self):
        connection = cx_Oracle.connect("system/69-Gilani-53@localhost:1521/orcl")
        return connection
    
    def InsertData(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            cid = self.courseId.get()
            cname = self.courseName.get().upper()
            cfees = self.courseFees.get()
            query = f"INSERT INTO COURSES (CID, CNAME, CFEES) VALUES ('{cid}', '{cname}', '{cfees}')"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            tmsg.showinfo("Good Job", f"Course # {self.courseId.get()} Is Inserted Successfully")
            self.FetchCourses()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Insertion Error", f"Check Course ID Or Error Due To {e}")
    
    def DeleteData(self):
        a = tmsg.askyesnocancel("Attention", "Do You Want To Delete The Course Details Permenantly?")
        if a == True:
            try:
                connection = self.connection()
                cursor = connection.cursor()
                query = f"DELETE FROM COURSES WHERE CID = '{self.courseId.get()}'"
                cursor.execute(query)
                connection.commit()
                cursor.close()
                tmsg.showinfo("Good Job", f"Course # {self.courseId.get()} Is Deleted Successfully")
                self.FetchCourses()
                self.Clear()
            except Exception as e:
                tmsg.showerror("Deletion Error", f"Check Course ID Or Error Due To {e}")
        else:
            pass
    
    def getCursor(self, event):
        cursor = self.Tree.focus()
        item = self.Tree.item(cursor)
        row = item['values']
        self.courseId.set(row[0])
        self.courseName.set(row[1])
        self.courseFees.set(row[2])

    def FetchCourses(self):
        connection = self.connection()
        cursor = connection.cursor()
        query = "SELECT * FROM COURSES ORDER BY CID ASC"
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
            query = f"UPDATE COURSES SET CNAME = '{self.courseName.get().upper()}', CFEES = '{self.courseFees.get()}' WHERE CID = '{self.courseId.get()}'"
            cursor.execute(query)
            connection.commit()
            cursor.close()
            tmsg.showinfo("Good Job", f"Course # {self.courseId.get()} Is Updated Successfully")
            self.FetchCourses()
            self.Clear()
        except Exception as e:
            tmsg.showerror("Updation Error", f"Check Course ID Or Error Due To {e}")
        
    def SearchCourse(self):
        try:
            connection = self.connection()
            cursor = connection.cursor()
            if self.SearchBy.get() == "Course ID":
                query = f"SELECT * FROM COURSES WHERE CID = '{self.Search.get()}'"
            elif self.SearchBy.get() == "Course Name":
                query = f"SELECT * FROM COURSES WHERE CNAME = '{self.Search.get().upper()}'"
            else:
                query = f"SELECT * FROM COURSES WHERE '{self.SearchBy.get()}' = '{self.Search.get().upper()}'"
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
        if self.SearchBy.get() == 'Course ID' or self.SearchBy.get() == "Course Name" or self.SearchBy.get() == "Gmail":
            self.SearchE.config(state = "normal")
        else:
            self.SearchE.config(state = "disabled")
        self.SearchE.after(500, self.ValidateSearchBar)

if __name__ == '__main__':
    root = Tk()
    Course = Course(root)
    root.mainloop()