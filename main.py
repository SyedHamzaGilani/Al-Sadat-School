from tkinter import *
from PIL import Image, ImageTk
import time
from datetime import date
from CertificateWindow import Certificate
from AttendenceWindow import Attendence
from EnrollWindow import Enroll
from CourseWindow import Course
from StudentWindow import Student
import ctypes


class MainWindow:

    def __init__(self, root):
        self.root = root
        self.trasImg = 0
        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()-70
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.root.geometry(f'{self.width}x{self.height}+0+0')
        self.root.resizable(False, False)
        self.root.title('Al Sadat Islamic School Of Computer Sciences')
        self.root.config(bg='Cyan')
        self.frame1 = Frame(self.root, bd=5, relief=SUNKEN)
        self.frame1.place(x=5, y=10, width=1355, height=50)
        # Label(self.frame1, text = "Welcome To Computer Science Of Scholars", font = "Helvatica 20 bold").pack(pady = 5)
        Label(self.frame1, text="Al Sadat Islamic School Of Computer Sciences", font="Helvatica 20 bold").pack(pady=5)

        ####################################

        self.frame2 = Frame(self.root, bd=5, relief=SUNKEN)
        image1 = Image.open("images/school.jpg")
        image1 = image1.resize((1366, 768-80), Image.Resampling.LANCZOS)
        self.image1 = ImageTk.PhotoImage(image1)
        self.ImgLabel = Label(self.frame2, image=self.image1)
        self.ImgLabel.place(x=0, y=0)
        self.ChangeImages()
        self.frame2.place(x=5, y=70, width=1355, height=560)

        ####################################

        self.frame3 = Frame(self.root, bd=5, relief=SUNKEN)
        self.frame3.place(x=5, y=640, width=1355, height=60)
        Button(self.frame3, text="Courses", font="Courier 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3",
               cursor="hand2", command=self.CourseWindow).place(x=5, y=5, height=45, width=130)
        Button(self.frame3, text="Students", font="Courier 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3",
               cursor="hand2", command=self.StudentWindow).place(x=140, y=5, height=45, width=130)
        Button(self.frame3, text="Enrollment", font="Courier 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3",
               cursor="hand2", command=self.EnrollWindow).place(x=275, y=5, height=45, width=150)
        Button(self.frame3, text="Attendence", font="Courier 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3",
               cursor="hand2", command=self.AttendenceWindow).place(x=430, y=5, height=45, width=160)
        Button(self.frame3, text="Certificate", font="Courier 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3",
               cursor="hand2", command=self.CertificateWindow).place(x=595, y=5, height=45, width=170)
        Button(self.frame3, text="Log Out", font="Courier 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3",
               cursor="hand2", command=self.root.destroy).place(x=770, y=5, height=45, width=160)
        self.TimeLabel = Label(self.frame3, font="Helvatica 14 bold")
        self.Time()
        self.TimeLabel.place(x=935, y=12)

    def ChangeImages(self):
        images = ['images/school.jpg', 'images/school1.jpg', 'images/school2.jpg', 'images/school3.jpg', 'images/school4.jpg', 'images/cert2.jpg', 'images/logo1.jpg', 'images/logo2.jpg']
        if self.trasImg >= len(images):
            self.trasImg = 0
        image = Image.open(f"{images[self.trasImg]}")
        image = image.resize((1343, 538), Image.Resampling.LANCZOS)
        self.image = ImageTk.PhotoImage(image)
        self.ImgLabel.config(image=self.image)
        self.ImgLabel.after(3000, self.ChangeImages)
        self.trasImg += 1

    def Time(self):
        today = date.today()
        today = today.strftime("%d / %B / %Y")
        string = time.strftime(f"%I : %M : %S  %p (%A) {today}")
        self.TimeLabel.config(text=string)
        self.TimeLabel.after(1000, self.Time)

    def StudentWindow(self):
        try:
            self.newwindow.destroy()
        except:
            pass
        self.newwindow = Toplevel(self.root)
        self.student = Student(self.newwindow)
        try:
            self.newwindow1.destroy()
            self.newwindow2.destroy()
            self.newwindow3.destroy()
            self.newwindow4.destroy()
        except:
            pass

    def CourseWindow(self):
        try:
            self.newwindow1.destroy()
        except:
            pass
        self.newwindow1 = Toplevel(self.root)
        self.course = Course(self.newwindow1)
        try:
            self.newwindow.destroy()
            self.newwindow2.destroy()
            self.newwindow3.destroy()
            self.newwindow4.destroy()
        except:
            pass

    def EnrollWindow(self):
        try:
            self.newwindow2.destroy()
        except:
            pass
        self.newwindow2 = Toplevel(self.root)
        self.enroll = Enroll(self.newwindow2)
        try:
            self.newwindow.destroy()
            self.newwindow1.destroy()
            self.newwindow3.destroy()
            self.newwindow4.destroy()
        except:
            pass

    def AttendenceWindow(self):
        try:
            self.newwindow3.destroy()
        except:
            pass
        self.newwindow3 = Toplevel(self.root)
        self.attendence = Attendence(self.newwindow3)

        try:
            self.newwindow.destroy()
            self.newwindow1.destroy()
            self.newwindow2.destroy()
            self.newwindow4.destroy()
        except:
            pass

    def CertificateWindow(self):
        try:
            self.newwindow4.destroy()
        except:
            pass
        self.newwindow4 = Toplevel(self.root)
        self.certificate = Certificate(self.newwindow4)

        try:
            self.newwindow.destroy()
            self.newwindow1.destroy()
            self.newwindow2.destroy()
            self.newwindow3.destroy()
        except:
            pass


if __name__ == '__main__':
    root = Tk()
    main = MainWindow(root)
    root.mainloop()
