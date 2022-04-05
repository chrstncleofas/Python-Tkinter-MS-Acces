from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
import pyodbc

class studentForm:
    def __init__(self, root):

        self.root = root
        self.root.title("Student Information System in Python Tkinter and Ms Access Database")
        self.root.withdraw()
        x = (self.root.winfo_screenwidth() -
             self.root.winfo_reqwidth()) / 10
        y = (self.root.winfo_screenheight() - self.root.winfo_reqheight()) / 30
        self.root.geometry("1130x665+%d+%d" % (x, y))

        # Head Label
        self.lbl_title = Label(
            self.root, text="STUDENT INFORMATION SYSTEM", bd=4, fg="white", relief=RIDGE,  bg="#fc5c00", font=("roboto sans-serif", 23), pady=7)
        self.lbl_title.pack(side=TOP, fill=X)

        # Text Box Input Frame 
        self.txt_frame = Frame(
            self.root, bd=4, relief=RIDGE, bg="navy")
        self.txt_frame.place(x=0, y=55, width=1130, height=250)

        # Frame Table
        self.detail_frame = Frame(
            self.root, bd=4, relief=RIDGE, bg="navy")
        self.detail_frame.place(x=0, y=300, width=1130, height=364)

        # Buttons Frame
        self.btn_frame = Frame(
            self.detail_frame, bd=3, relief=RIDGE, bg="navy")
        self.btn_frame.place(x=0, y=260, width=1122, height=98)

        # Button Here
        # Space for Buttons
        self.space_label = Label(self.btn_frame, text='',
                                 font=('', 10), bg="navy", fg="white").grid(row=1, column=2, pady=20, padx=5)
        # Save
        self.save_btn = Button(self.btn_frame, text='INSERT', width=24, height=2, bd=2, relief=FLAT, bg="#fc5c01", fg="white",
                               font=("roboto sans-serif", 13, "bold"), command=self.insert_item)
        self.save_btn.grid(row=1, column=3, pady=20, padx=10)
        # Update
        self.update_btn = Button(self.btn_frame, text='UPDATE', width=24, height=2, bd=2, relief=FLAT, bg="#fc5c01", fg="white",
                                 font=("roboto sans-serif", 13, "bold"), command=self.update_item)
        self.update_btn.grid(row=1, column=4, pady=20, padx=10)
        # Delete
        self.delete_btn = Button(self.btn_frame, text='DELETE', width=24, height=2, bg="#fc5c01", fg="white", bd=2, relief=FLAT,
                                 font=("roboto sans-serif", 13, "bold"), command=self.delete_item)
        self.delete_btn.grid(row=1, column=5, padx=12)
        # Clear
        self.clear_btn = Button(self.btn_frame, text='CLEAR', width=24, height=2, bg="#fc5c01", fg="white", bd=2, relief=FLAT,
                                font=("roboto sans-serif", 13, "bold"), command=self.clear_text)
        self.clear_btn.grid(row=1, column=6, padx=12)

        # Fields Textbox
        # Space for Fields
        self.space_label = Label(self.txt_frame, text='',
                                 font=('', 10), bg="navy", fg="white").grid(row=0, column=2, pady=10, sticky=W, padx=43)

        # Student ID
        self.student = StringVar()
        self.student_label = Label(self.txt_frame, text='STUDENT ID :',
                                   font=('bold', 14), bg="navy", fg="white").grid(row=1, column=0, sticky=W, padx=52)
        self.student_entry = Entry(
            self.txt_frame, textvariable=self.student, width=25, bd=3, font=("bold", 15))
        self.student_entry.grid(row=1, column=1)

        # First Name
        self.fname = StringVar()
        self.fname_label = Label(self.txt_frame, text='FIRST NAME :', font=(
            'bold', 14), bg="navy", fg="white").grid(row=2, column=0, sticky=W, padx=50)
        self.fname_entry = Entry(self.txt_frame, textvariable=self.fname,
                                 width=25, bd=3, font=("bold", 15))
        self.fname_entry.grid(row=2, column=1)

        # Middle Name
        self.mname = StringVar()
        self.mname_label = Label(self.txt_frame, text='MIDDLE NAME :', font=('bold', 14),
                                 bg="navy", fg="white").grid(row=3, column=0, sticky=W, padx=50)
        self.mname_entry = Entry(self.txt_frame, textvariable=self.mname,
                                 width=25, bd=3, font=("bold", 15))
        self.mname_entry.grid(row=3, column=1)

        # Last Name
        self.lname = StringVar()
        self.lname_label = Label(self.txt_frame, text='LAST NAME :', font=(
            'bold', 14), bg="navy", fg="white").grid(row=1, column=2, sticky=W, padx=50)
        self.lname_entry = Entry(self.txt_frame, textvariable=self.lname,
                                 width=25, bd=3, font=("bold", 15))
        self.lname_entry.grid(row=1, column=3)

        # Course
        self.course = StringVar()
        self.course_label = Label(self.txt_frame, text='COURSE :', font=('bold', 14),
                                  bg="navy", fg="white").grid(row=2, column=2, sticky=W, padx=50, pady=20)
        self.course_entry = Entry(self.txt_frame, textvariable=self.course,
                                  width=25, bd=3, font=("bold", 15))
        self.course_entry.grid(row=2, column=3)

        # Data Table in TreeView
        # Listbox Frame
        self.list_frame = Frame(
            self.detail_frame, bd=2, relief=RIDGE, bg="navy")
        self.list_frame.place(x=0, y=0, width=1122, height=265)

        # Treeview Scrollbar
        scroll_x = Scrollbar(self.list_frame, orient=HORIZONTAL)
        scroll_y = Scrollbar(self.list_frame, orient=VERTICAL)

        # Treeview
        design = ttk.Style()
        # design.configure("Treeview")
        design.theme_use("clam")

        self.data_list = ttk.Treeview(self.list_frame, height=12, columns=("StudentID", "Fname", "Mname", "Lname", "Course"), xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)

        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)

        self.data_list.configure(yscrollcommand=scroll_x.set)
        scroll_x.configure(command=self.data_list.xview)

        self.data_list.configure(yscrollcommand=scroll_y.set)
        scroll_y.configure(command=self.data_list.yview)

        self.data_list.heading("StudentID", text="STUDENT ID")
        self.data_list.heading("Fname", text="FIRST NAME")
        self.data_list.heading("Mname", text="MIDDLE NAME")
        self.data_list.heading("Lname", text="LAST NAME")
        self.data_list.heading("Course", text="COURSE")

        self.data_list['show'] = 'headings'

        self.data_list.column("StudentID", width=80, anchor=tk.CENTER)
        self.data_list.column("Fname", width=160, anchor=tk.CENTER)
        self.data_list.column("Mname", width=146, anchor=tk.CENTER)
        self.data_list.column("Lname", width=140, anchor=tk.CENTER)
        self.data_list.column("Course", width=200, anchor=tk.CENTER)
        self.data_list.pack(fill=BOTH, expand=1)
        self.data_list.bind('<ButtonRelease-1>', self.select_item)
        # End of Design Treeview 

        # To show all Data in Treeview
        self.dataList()

    # Data Table 
    def dataList(self):
        try:
            con = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\dbFolder\pydb.mdb;')
            cur = con.cursor()
            cur.execute("SELECT * FROM tblstudent")
            rows = cur.fetchall()
            if len(rows) != 0:
                self.data_list.delete(*self.data_list.get_children())
                for row in rows:
                    self.data_list.insert('', END, values=(row[0], row[1], row[2], row[3], row[4]))
                con.commit()
            con.close()
        except pyodbc.Error as e:
            messagebox.showerror("Error", e, "Error Message")

    # Save Function
    def insert_item(self):
        try :
            if self.student_entry.get() == '' or self.fname_entry.get() == '' or self.mname_entry.get() == '' or self.lname_entry.get() == '' or self.course_entry.get() == '':
                messagebox.showerror(
                    'Required Fields', 'Please include all fields')
                return
            else:
                con = pyodbc.connect(
                    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\dbFolder\pydb.mdb;')
                cur = con.cursor()

                cur.execute("INSERT INTO tblstudent([StudentID], [Fname], [Mname], [Lname],[Course]) VALUES('"+(self.student.get())+"','"+(self.fname.get())+"','"+(self.mname.get())+"','"+(self.lname.get())+"','"+(self.course.get())+"')")

                con.commit()
                con.close()
                messagebox.showinfo("Success", "Saved into database.....")
            self.clear_text()
            self.dataList()
        except pyodbc.Error as e:
            messagebox.showerror("Error", e, "Error Message")

    # Update Function
    def update_item(self):
        try:
            con = pyodbc.connect(
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\dbFolder\pydb.mdb;')
            cur = con.cursor()

            cur.execute(
            "UPDATE tblstudent SET [Fname]=?, [Mname]=?, [Lname]=?, [Course]=? WHERE [StudentID]=?",
            (self.fname_entry.get(),
             self.mname_entry.get(),
             self.lname_entry.get(),
             self.course_entry.get(),
             self.student_entry.get()
             ))
            con.commit()
            messagebox.showinfo("Success", "Update Successfuly")
            self.dataList()
            self.clear_text()
            con.close()
        except pyodbc.Error as e:
            messagebox.showerror("Error", e, "Error Message")

    # Delete Funtion
    def delete_item(self):
        try:
            con = pyodbc.connect(
                r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\dbFolder\pydb.mdb;')
            cur = con.cursor()

            cur.execute("DELETE FROM tblstudent WHERE StudentID=?",
                            self.student.get())

            con.commit()   
            messagebox.askyesno('Delete', 'Do you want to delete this file')
            select_item = self.data_list.selection()[0]
            self.data_list.delete(select_item)
            self.clear_text()
            con.close()
        except pyodbc.Error as e:
            messagebox.showerror("Error", e, "Error Message")

    # Clear Items
    def clear_text(self):
        self.student_entry.delete(0, END)
        self.fname_entry.delete(0, END)
        self.mname_entry.delete(0, END)
        self.lname_entry.delete(0, END)
        self.course_entry.delete(0, END)

    # Select in Treeview Data
    def select_item(self, ev):
        cursor_row = self.data_list.focus()
        contents = self.data_list.item(cursor_row)
        row = contents['values']
        self.student.set(row[0])
        self.fname.set(row[1])
        self.mname.set(row[2])
        self.lname.set(row[3])
        self.course.set(row[4])
        

root = Tk()
obj = studentForm(root)
root.deiconify()
root.mainloop()