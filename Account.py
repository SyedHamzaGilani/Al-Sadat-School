from tkinter import *
import cx_Oracle
import tkinter.messagebox as tmsg
import ctypes
class Account:
    def __init__(self, root):
        self.root = root
        self.root.geometry("500x500+400+100")
        self.root.resizable(False, False)
        self.c = 0
        self.root.title("Account")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.root.config(bg = "Black")
        frame = Frame(self.root, bg = "#f0f0f0")
        frame.place(x = 5, y = 5, width = 490, height = 490)
        self.createUser = StringVar()
        self.oldPass = StringVar()
        self.newPass = StringVar()
        self.label = Label(frame, text = "Create Login Details", font = "Courier 20 bold underline", bg = "#f0f0f0", fg = "Black")
        self.label.place(x = 90, y = 10)
        Label(frame, text = "User Name", font = "Arial 13 bold", bg = "#f0f0f0", fg = "Black").place(x = 20+50, y = 90)
        Entry(frame, textvariable = self.createUser,font = "Arial 12 bold", fg = "Black").place(x = 160+50, y = 90)
        Label(frame, text = "Old Password", font = "Arial 13 bold", bg = "#f0f0f0", fg = "Black").place(x = 20+50, y = 160)
        Entry(frame, textvariable=self.oldPass, font="Arial 12 bold", fg="Black").place(x=160+50, y=160)
        Label(frame, text = "New Password", font = "Arial 13 bold", bg = "#f0f0f0", fg = "Black").place(x = 20+50, y = 230)
        Entry(frame, textvariable=self.newPass, font="Arial 12 bold", fg="Black").place(x=160+50, y=230)
        Button(frame, text = "Exit", font = "Courier 14 bold", bg = "#b3b3b3", fg = "Black", cursor = "hand2", command = self.root.destroy).place(x = 50+30, y = 320, width = 95)
        Button(frame, text = "Clear", font = "Courier 14 bold", bg = "#b3b3b3", fg = "Black", cursor = "hand2", command = self.Clear).place(x = 150+30, y = 320, width = 95)
        Button(frame, text = "Reset Now", font = "Courier 14 bold", bg = "#b3b3b3", fg = "Black", cursor = "hand2", command = self.new).place(x = 250+30, y = 320)


    def Clear(self):
        self.createUser.set("")
        self.oldPass.set("")
        self.newPass.set("")

    def new(self):
        self.Fetch()
        # self.newAccount(self.createUser.get(), self.newPass.get())

    def Connection(self):
        connection = cx_Oracle.connect("system/69-Gilani-53@localhost:1521/orcl")
        return connection

    def Fetch(self):
        try:
            connection = self.Connection()
            cursor = connection.cursor()
            query = "SELECT * FROM LOGIN"
            cursor.execute(query)
            rows = cursor.fetchone()
            a = str(self.createUser.get().upper())
            b = str(self.oldPass.get())
            try:
                if a == rows[0] and b == rows[1]:
                    self.newAccount(self.createUser.get().upper(), self.newPass.get())
                    tmsg.showinfo("Successful", "Password Updated Successfully!")
                    self.root.destroy()
                else:
                    tmsg.showerror("Error", "Operation Denied!, Error In Username Or Old Password")
                    self.root.destroy()
            except Exception as er:
                tmsg.showerror("Error", f"Error due to {er}")
            connection.commit()
            cursor.close()
        except Exception as er:
            tmsg.showerror("Error", f"Error due to {er}")

    def newAccount(self, user, passw):
        connection = self.Connection()
        cursor = connection.cursor()
        query = f"UPDATE LOGIN SET PASSWORD = '{passw}' WHERE USERNAME = '{user}'"
        cursor.execute(query)
        connection.commit()
        cursor.close()
        self.Clear()


if __name__ == "__main__":
    root = Tk()
    account = Account(root)
    root.mainloop()