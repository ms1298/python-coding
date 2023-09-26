import tkinter
from tkinter import *
from tkcalendar import *
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl


window = tkinter.Tk()
window.title("Data Entry Form") 
frame = tkinter.Frame(window)
frame.pack()


#Operation Section



def enter_data():
    accepted= accept_var.get()

    if accepted == "Accepted":


        firstName =  first_name_Entry.get()
        lastName = last_name_Entry.get()
        

        if firstName and lastName:
                
            titleName = title_combobox.get()
            age = dob_Entry.get()
            
            mob_No = mob_Entry.get()
            image = img_Entry.get()
            gender_get = radio_var.get()

            
            print("Data Entry done!")

            # course about

            registraion_status = reg_status_var.get()
            course_pref = course_combo2.get()
            print("---------------")
            courseName = course_combo1.get()

            filepath = "E:\Some_Project\Octagon.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title","Date of Birth",
                            "Mobile No.", "Course Name",
                            "Image", "Gender", "Registration Status","Course Preference"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstName, lastName, titleName, age, mob_No, courseName, 
                        image, gender_get, registraion_status,course_pref])
            workbook.save(filepath)
                

        else:
            tkinter.messagebox.showwarning(title="Error", message="Please Enter first name & last name.")
    else:
        tkinter.messagebox.showwarning(title="Error! ", message="You don't have check the 'terms & conditions' " )








#Design Section
#saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_Entry = tkinter.Entry(user_info_frame)
last_name_Entry = tkinter.Entry(user_info_frame)
first_name_Entry.grid(row=1, column=0)
last_name_Entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["","Md.", "Md", "Dr.","Eng.", "Mr.", "Mrs."])
title_label.grid(row=0,column=2)
title_combobox.grid(row=1,column=2)

age_label = tkinter.Label(user_info_frame,text="Date of Birth")
dob_Entry = tkinter.Entry(user_info_frame)
age_label.grid(row=2, column=0)
dob_Entry.grid(row=2, column=1)


#date--------------------------------

def new_window_for_DOB():
    global cal
    root = Tk()

    root.geometry("350x380")
    cal = Calendar(root, selectmode = 'day',
                year = 2020, month = 5,
                day = 22)

    cal.pack(pady=5)

    def grad_date():

        date.config(text="" +(cal.get_date()))

    # Add Button and Label
    Button(root, text = "Get Date",
        command = grad_date).pack(pady = 20)

    date = (Label(user_info_frame))
    date.grid(row=2, column=1)

    grad_date()

    root.mainloop()
    root.destroy()

    

dob_btn = Button(user_info_frame, text="Select Date", command= new_window_for_DOB)
dob_btn.grid(row=2,column=2)

#date -----------end----------------







mob_label = tkinter.Label(user_info_frame, text="Mobile No.")
mob_label.grid(row=4, column=0)
mob_Entry = tkinter.Entry(user_info_frame)
mob_Entry.grid(row=4, column=1)

course_lable = tkinter.Label(user_info_frame, text="Course Name")
course_lable.grid(row=5, column=0)

course_combo1 = ttk.Combobox(user_info_frame, values=["JavaScript", "AngularJS", "Node.js", "Python", 
                                                        "Linux", "Java", "SAS", "C++", "C#", "R","MS Word",
                                                        "MS Excel","MS Power Point","Video Editing","Autocad",
                                                        "Animation","Graphics Design"])
course_combo1.grid(row=5,column=1)


img_label = tkinter.Label(user_info_frame, text="Upload Image")
img_label.grid(row=6, column=0)
img_Entry = tkinter.Entry(user_info_frame)
img_Entry.grid(row=6,column=1)

gender_laberl = tkinter.Label(user_info_frame, text="Gender")
gender_laberl.grid(row=7, column=0)

def gender():
    global radio_var
    radio_var = tkinter.StringVar(value='Gender not Selectd')

    r1 = tkinter.Radiobutton(user_info_frame, text='Male', variable=radio_var, value = 'Male', offvalue =None )
    r1.grid(row=7, column=1)
    r2 = tkinter.Radiobutton(user_info_frame, text='Female', variable= radio_var, value='Female', offvalue=None)
    r2.grid(row=7, column=2)
    r3 = tkinter.Radiobutton(user_info_frame, text='Others',variable= radio_var, value='Others', offvalue=None)
    r3.grid(row=7, column=3)

gender()


for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)


#Saving Course Info
#registration

reg_status_var = tkinter.StringVar(value="Not registerd")

course_frame = tkinter.LabelFrame(frame, text="Registration and Course Status")
course_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)
registration_check = tkinter.Checkbutton(course_frame, text= "Currently Registered", 
                                         variable=reg_status_var, onvalue= "Registered", offvalue="Not registered"  )
registration_check.grid(row=1,column=0)



course_com = tkinter.Label(course_frame, text="Completed Course(s)")
numcouse_spinbox = ttk.Spinbox(course_frame, from_=0, to="infinity")
course_com.grid(row=0,column=1)
numcouse_spinbox.grid(row=1, column=1)

course_pref_label = tkinter.Label(course_frame, text="Course Preference")
course_combo2 = ttk.Combobox(course_frame, values=["JavaScript", "AngularJS", "Node.js", "Python", 
                                                    "Linux", "Java", "SAS", "C++", "C#", "R","MS Word",
                                                    "MS Excel","MS Power Point","Video Editing","Autocad",
                                                    "Animation","Graphics Design"])

course_pref_label.grid(row=0,column=2)
course_combo2.grid(row=1, column=2)

for widget in course_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#Accept term
accept_var = tkinter.StringVar(value="Not Accepted")
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditons")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms & conditons.", 
                                  variable= accept_var, onvalue="Accepted", offvalue="Not Accepted" )
terms_check.grid(row=0, column=0)


#Button
btn_insert = tkinter.Button(terms_frame, text="Insert data", command=enter_data)
btn_insert.grid(row=3, column=0 ,padx=20, pady=10 )

btn_edit = tkinter.Button(terms_frame, text="Edit data" )
btn_edit.grid(row=3, column=1, padx=20, pady=10)

btn_show = tkinter.Button(terms_frame, text="Show data" )
btn_show.grid(row=3, column=2, padx=20, pady=10)

btn_delete = tkinter.Button(terms_frame, text="Delete  data")
btn_delete.grid(row=3, column=3, padx=20, pady=12)

for widget in terms_frame.winfo_children():
    widget.grid_configure(padx=10, pady=10)





window.mainloop()