from datetime import date
from PIL import Image, ImageDraw, ImageFont
from tkinter import *
from PIL import Image, ImageTk
import tkinter.messagebox as tmsg
import os
import cx_Oracle
from tkinter import ttk as tk
import ctypes


class Certificate:
    def __init__(self, root):
        self.root = root
        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()-70
        self.root.geometry(f'{self.width}x{self.height}+0+0')
        self.root.resizable(False, False)
        self.root.title('Certificate Details')
        self.root.config(bg = "Cyan")
        myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.root.iconbitmap(default='images/icon2.ico')
        self.Name = "Syed Hamza Gilani"
        self.CName = "Data Science"
        self.date = self.GetDate()
        self.frame1 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame1.place(x = 5, y = 5, width = 1355, height = 150)
        self.SID = StringVar()
        self.CID = StringVar()
        Label(self.frame1, text = "S-ID", font = "helvatica 12 bold").place(x = 5, y = 15)
        Label(self.frame1, text = "C-ID", font = "helvatica 12 bold").place(x = 160, y = 15)
        connection = self.Connection()
        cursor = connection.cursor()
        cursor1 = connection.cursor()
        query = "SELECT DISTINCT(SID) FROM ENROLLEDCOURSE ORDER BY SID ASC"
        cursor.execute(query)
        rows = cursor.fetchall()
        connection.commit()
        cursor.close()
        sidE = tk.Combobox(self.frame1, textvariable=self.SID, font = "helvatica 11 bold")
        sidE['values'] = rows
        sidE.place(x = 50, y = 15, width = 100)

        query1 = "SELECT DISTINCT(CID) FROM ENROLLEDCOURSE ORDER BY CID"
        cursor1.execute(query1)
        rows1 = cursor1.fetchall()
        connection.commit()
        cursor1.close()
        cidE = tk.Combobox(self.frame1, textvariable=self.CID, font = "helvatica 11 bold")
        cidE['values'] = rows1
        cidE.place(x = 210, y = 15, width = 130)
        Button(self.frame1, text="Generate", font="Helvatica 12 bold", bd=3, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.getDetails).place(x = 420, y = 15, height=25, width=220)
        Label(self.frame1, text = "Left Right Margin Of Name", font = "Helvatica 12 bold").place(x = 5, y = 20+40)
        self.sliderNameX = Scale(self.frame1, from_=0, to=2600, orient=HORIZONTAL)
        self.sliderNameX.place(x = 230, y = 5+40)
        Label(self.frame1, text = "Top Bottom Margin Of Name", font = "Helvatica 12 bold").place(x = 5, y = 70+40)
        self.sliderNameY = Scale(self.frame1, from_=0, to=1400, orient=HORIZONTAL)
        self.sliderNameY.place(x = 230, y = 55+40)
        Label(self.frame1, text = "Left Right Margin Of Course", font = "Helvatica 12 bold").place(x = 405, y = 20+40)
        self.sliderCourseX = Scale(self.frame1, from_=0, to=2600, orient=HORIZONTAL)
        self.sliderCourseX.place(x = 650, y = 5+40)
        Label(self.frame1, text = "Top Bottom Margin Of Course", font = "Helvatica 12 bold").place(x = 405, y = 70+40)
        self.sliderCourseY = Scale(self.frame1, from_=0, to=1400, orient=HORIZONTAL)
        self.sliderCourseY.place(x = 650, y = 55+40)
        self.sliderNameX.set('835')
        self.sliderNameY.set('835')
        self.sliderCourseX.set('920')
        self.sliderCourseY.set('1010')


        Button(self.frame1, text="Update Text", font="Helvatica 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.UpdateText).place(x = 800, y = 20, height=45, width=190)
        Button(self.frame1, text="Reset", font="Helvatica 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.Reset).place(x = 800, y = 70, height=45, width=190)
        Button(self.frame1, text="Save Image", font="Helvatica 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.SaveImg).place(x = 1000, y = 20, height=45, width=190)
        Button(self.frame1, text="Exit Now", font="Helvatica 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.root.destroy).place(x = 1000, y = 70, height=45, width=190)
        Button(self.frame1, text="Print Now", font="Helvatica 18 bold", bd=3, relief=GROOVE, bg="#b3b3b3", cursor="hand2", command=self.InsertCert).place(x = 1200, y = 20, height=95, width=140)
        self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
        self.frame2.place(x = 5, y = 160, width=self.width-10, height=self.height-170)
        self.name = "Raw"
        self.x = 835
        self.y = 835
        self.x1 = 920
        self.y1 = 1010
        font = ImageFont.truetype('calibri.ttf',70)
        font1 = ImageFont.truetype('calibri.ttf',50)
        self.img1 = Image.open('images/Certificate-2.jpg')
        cert = ImageDraw.Draw(self.img1)
        cert.text(xy=(950,520),text=f"Cert No : {self.getCertNo()}",fill=(0, 0, 0),font=font1, stroke_width=2, stroke_fill="black")
        name = ImageDraw.Draw(self.img1)
        name.text(xy=(self.x,self.y),text=f"{self.Name}",fill=(0, 0, 0),font=font, stroke_width=2, stroke_fill="black")
        course = ImageDraw.Draw(self.img1)
        course.text(xy = (self.x1, self.y1), text = f"{self.CName}", fill=(0, 0, 0),font=font, stroke_width=2, stroke_fill="black")
        date = ImageDraw.Draw(self.img1)
        date.text(xy = (1030, 1140), text = f"{self.date}", fill=(0, 0, 0),font=font1, stroke_width=1, stroke_fill="black")
        self.img1.save(f'images/Raw.jpg')
        image = Image.open('images/Raw.jpg')
        image = image.resize((self.width-10, self.height-170), Image.Resampling.LANCZOS)
        self.img = ImageTk.PhotoImage(image)
        Label(self.frame2, image = self.img).place(x = 0, y = 0)

    def ShowImage(self):
        try:
            self.frame1.state('destroy')
        except:
            self.frame2 = Frame(self.root, bd = 5, relief = SUNKEN)
            self.frame2.place(x = 5, y = 160, width=self.width-10, height=self.height-170)
            image = Image.open('images/Raw.jpg')
            image = image.resize((self.width-10, self.height-170), Image.Resampling.LANCZOS)
            self.img = ImageTk.PhotoImage(image)
            Label(self.frame2, image = self.img).place(x = 0, y = 0)

    def DuplicateCert(self):
        x = self.sliderNameX.get()
        y = self.sliderNameY.get()
        x1 = self.sliderCourseX.get()
        y1 = self.sliderCourseY.get()
        self.x = x
        self.y = y
        self.x1 = x1
        self.y1 = y1
        font = ImageFont.truetype('calibri.ttf',70)
        font1 = ImageFont.truetype('calibri.ttf',50)
        self.img1 = Image.open('images/Certificate-2.jpg')
        cert = ImageDraw.Draw(self.img1)
        cert.text(xy=(950,520),text=f"Cert No : {self.getCertNo()}",fill=(0, 0, 0),font=font1, stroke_width=2, stroke_fill="black")
        name = ImageDraw.Draw(self.img1)
        name.text(xy=(self.x,self.y),text=f"{self.Name}",fill=(0, 0, 0),font=font, stroke_width=2, stroke_fill="black")
        Duplicate = ImageDraw.Draw(self.img1)
        Duplicate.text(xy=(1555,625),text=f"( Duplicate )",fill=(255, 0, 0),font=font, stroke_width=2, stroke_fill="black")
        course = ImageDraw.Draw(self.img1)
        course.text(xy = (self.x1, self.y1), text = f"{self.CName}", fill=(0, 0, 0),font=font, stroke_width=2, stroke_fill="black")
        date = ImageDraw.Draw(self.img1)
        date.text(xy = (1030, 1140), text = f"{self.date}", fill=(0, 0, 0),font=font1, stroke_width=1, stroke_fill="black")
        self.img1.save(f'images/Raw.jpg')
        self.ShowImage()
        self.Print()


    def UpdateText(self):
        x = self.sliderNameX.get()
        y = self.sliderNameY.get()
        x1 = self.sliderCourseX.get()
        y1 = self.sliderCourseY.get()
        self.x = x
        self.y = y
        self.x1 = x1
        self.y1 = y1
        font = ImageFont.truetype('calibri.ttf',70)
        font1 = ImageFont.truetype('calibri.ttf',50)
        self.img1 = Image.open('images/Certificate-2.jpg')
        cert = ImageDraw.Draw(self.img1)
        cert.text(xy=(950,520),text=f"Cert No : {self.getCertNo()}",fill=(0, 0, 0),font=font1, stroke_width=2, stroke_fill="black")
        name = ImageDraw.Draw(self.img1)
        name.text(xy=(self.x,self.y),text=f"{self.Name}",fill=(0, 0, 0),font=font, stroke_width=2, stroke_fill="black")
        course = ImageDraw.Draw(self.img1)
        course.text(xy = (self.x1, self.y1), text = f"{self.CName}", fill=(0, 0, 0),font=font, stroke_width=2, stroke_fill="black")
        date = ImageDraw.Draw(self.img1)
        date.text(xy = (1030, 1140), text = f"{self.date}", fill=(0, 0, 0),font=font1, stroke_width=1, stroke_fill="black")
        self.img1.save(f'images/Raw.jpg')
        self.ShowImage()
    
    def Reset(self):
        self.sliderNameX.set('835')
        self.sliderNameY.set('835')
        self.sliderCourseX.set('920')
        self.sliderCourseY.set('1010')
        self.UpdateText()
    
    def SaveImg(self):
        try:
            self.img1.save(f'images/Raw.jpg')
            tmsg.showinfo("Good Job", "Image Is Sucessfully Saved !")
        except Exception as e:
            tmsg.showerror("Error", f"Error Due To {e}")

    def Print(self):
        filename = r'E:\SchoolManagementSystem\images\Raw.jpg'
        os.startfile(filename, "print")

    def Connection(self):
        connection = cx_Oracle.connect('system/69-Gilani-53@localhost:1521/orcl')
        return connection
    
    def getDetails(self):
        connection = self.Connection()
        try:
            cursor1 = connection.cursor()
            cursor2 = connection.cursor()
            query = f"SELECT CNAME FROM COURSES WHERE CID = '{self.CID.get()}'"
            query2 = f"SELECT SNAME FROM STUDENTS WHERE SID = '{self.SID.get()}'"
            cursor1.execute(query)
            cursor2.execute(query2)
            row = cursor1.fetchone()
            row2 = cursor2.fetchone()
            for i in row:
                self.CName = i
            for j in row2:
                self.Name = j
            cursor1.close()
            connection.commit()
            self.UpdateText()
        except Exception as e:
            tmsg.showerror('Error', f'Error Due To {e}')
    
    def GetDate(self):
        today = date.today()
        today = today.strftime("%d / %B / %Y")
        return today
    
    def getCertNo(self):
        connection = self.Connection()
        cursor = connection.cursor()
        query = f"SELECT MAX(CERT_NO) AS CERT_NO FROM CERTIFICATE"
        cursor.execute(query)
        row = cursor.fetchone()
        if row[0] == None:
            self.CNo = "1"
            return self.CNo
        try:
            for i in row:
                self.CNo = int(i)+1
                if i ==  '':
                    self.CNo = "1"
            return self.CNo
        except:
            pass

    def InsertCert(self):
        connection = self.Connection()
        try:
            cursor = connection.cursor()
            query = f"SELECT * FROM ENROLLEDCOURSE WHERE CID = '{self.CID.get()}' AND SID = '{self.SID.get()}'"    
            cursor.execute(query)
            rows = cursor.fetchone()
            cursor.close()
            if rows == None:
                tmsg.showerror("Certificate Error", f"Student {self.SID.get()} Is Not Enrolled In Course {self.CID.get()}")
            else:
                try:
                    cursor1 = connection.cursor()
                    query1 = f"INSERT INTO CERTIFICATE (CERT_NO, CID, SID, CDATE) VALUES ('{self.CNo}', '{self.CID.get()}', '{self.SID.get()}', TO_DATE('{self.date}', 'DD / MONTH / YYYY'))"
                    cursor1.execute(query1)
                    cursor1.close()
                    connection.commit()
                    tmsg.showinfo("Successful", "Certificate Is Assigned Successfully")
                    self.Print()
                except Exception as e:
                    self.DuplicateCert()
        except:
            tmsg.showerror("Certificate Error", f"The Student {self.SID.get()} Is Not Enrolled In Course {self.CID.get()}")

        

if __name__ == '__main__':
    root = Tk()
    certificate = Certificate(root)
    root.mainloop()