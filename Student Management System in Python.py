#! /usr/bin/python3
#! /usr/bin/env python3

from tkinter import *
from tkinter.ttk import *
from tkinter import ttk
from tkcalendar import Calendar, DateEntry
import tkinter as tk
from datetime import date
from tkinter import messagebox
from tkinter import scrolledtext
import datetime






allstudents = {}














def getdatee():
    def print_sel():
        global selecteddate
        selecteddate = str(cal.selection_get()) + " ( " + str(calculateAge(cal.selection_get())) + " years old )"
        birthday1.configure(text= selecteddate)

    top = tk.Toplevel(window)

    cal = Calendar(top,
                   font="Arial 14", selectmode='day',
                   cursor="hand1", year=2018, month=2, day=5)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()



selecteddate = 1



def calculateAge(birthDate): 
    today = date.today() 
    age = today.year - birthDate.year 
  
    return age 





window = Tk()
window.title("Student Management System in Python")
window.geometry('1000x800')

firstheader = Label(window, text="Add a new student:", font=("Arial Bold", 15))
firstheader.grid(column=0, row=0)




stname = Label(window, text="First name: ")
stname.grid(column=1, row=2)
stnametxt = Entry(window,width=20)
stnametxt.grid(column=2, row=2)


stlname = Label(window, text="Last name: ")
stlname.grid(column=3, row=2)
stlnametxt = Entry(window,width=20)
stlnametxt.grid(column=4, row=2)




birthday = Label(window, text="Birthday : ")
birthday.grid(column=1, row=5)

getdatebtn = Button(window, text="Calender", command=getdatee)
getdatebtn.grid(column=2, row=5)

birthday1 = Label(window, text=" ")
birthday1.grid(column=4, row=5)







secheader = Label(window, text="Lesson :")
secheader.grid(column=1, row=6)
thirdheader = Label(window, text="Score :")
thirdheader.grid(column=2, row=6)
frthheader = Label(window, text="Average :")
frthheader.grid(column=4, row=6)



 
Art_state = BooleanVar()
Art = Checkbutton(window, text='Art', var=Art_state)
Art.grid(column=1, row=7)
Artcombo = Combobox(window)
Artcombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
Artcombo.grid(column=2, row=7)
Artcombo.current(19)
Average_of_Art1 = Label(window, text="Average_of_Art: ")
Average_of_Art1.grid(column=4, row=7)
Average_of_Art2 = Label(window, text="-")
Average_of_Art2.grid(column=5, row=7)


 
Mathematics_state = BooleanVar()
Mathematics = Checkbutton(window, text='Mathematics', var=Mathematics_state)
Mathematics.grid(column=1, row=8)
Mathematicscombo = Combobox(window)
Mathematicscombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
Mathematicscombo.grid(column=2, row=8)
Mathematicscombo.current(19)
Average_of_Mathematics1 = Label(window, text="Average_of_Mathematics: ")
Average_of_Mathematics1.grid(column=4, row=8)
Average_of_Mathematics2 = Label(window, text="-")
Average_of_Mathematics2.grid(column=5, row=8)



Music_state = BooleanVar()
Music = Checkbutton(window, text='Music', var=Music_state)
Music.grid(column=1, row=9)
Musiccombo = Combobox(window)
Musiccombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
Musiccombo.grid(column=2, row=9)
Musiccombo.current(19)
Average_of_Music1 = Label(window, text="Average_of_Music: ")
Average_of_Music1.grid(column=4, row=9)
Average_of_Music2 = Label(window, text="-")
Average_of_Music2.grid(column=5, row=9)

 
 
Dance_state = BooleanVar()
Dance = Checkbutton(window, text='Dance', var=Dance_state)
Dance.grid(column=1, row=10)
Dancecombo = Combobox(window)
Dancecombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
Dancecombo.grid(column=2, row=10)
Dancecombo.current(19)
Average_of_Dance1 = Label(window, text="Average_of_Dance: ")
Average_of_Dance1.grid(column=4, row=10)
Average_of_Dance2 = Label(window, text="-")
Average_of_Dance2.grid(column=5, row=10)


 
PhysicalScience_state = BooleanVar()
PhysicalScience = Checkbutton(window, text='Physical-Science', var=PhysicalScience_state)
PhysicalScience.grid(column=1, row=11)
PhysicalSciencecombo = Combobox(window)
PhysicalSciencecombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
PhysicalSciencecombo.grid(column=2, row=11)
PhysicalSciencecombo.current(19)
Average_of_PhysicalScience1 = Label(window, text="Average_of_Physical Science: ")
Average_of_PhysicalScience1.grid(column=4, row=11)
Average_of_PhysicalScience2 = Label(window, text="-")
Average_of_PhysicalScience2.grid(column=5, row=11)



EnglishLiterature_state = BooleanVar()
EnglishLiterature = Checkbutton(window, text='English-Literature', var=EnglishLiterature_state)
EnglishLiterature.grid(column=1, row=12)
EnglishLiteraturecombo = Combobox(window)
EnglishLiteraturecombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
EnglishLiteraturecombo.grid(column=2, row=12)
EnglishLiteraturecombo.current(19)
Average_of_EnglishLiterature1 = Label(window, text="Average_of_English Literature: ")
Average_of_EnglishLiterature1.grid(column=4, row=12)
Average_of_EnglishLiterature2 = Label(window, text="-")
Average_of_EnglishLiterature2.grid(column=5, row=12)



Chemistry_state = BooleanVar()
Chemistry = Checkbutton(window, text='Chemistry', var=Chemistry_state)
Chemistry.grid(column=1, row=13)
Chemistrycombo = Combobox(window)
Chemistrycombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
Chemistrycombo.grid(column=2, row=13)
Chemistrycombo.current(19)
Average_of_Chemistry1 = Label(window, text="Average_of_Chemistry: ")
Average_of_Chemistry1.grid(column=4, row=13)
Average_of_Chemistry2 = Label(window, text="-")
Average_of_Chemistry2.grid(column=5, row=13)



French_state = BooleanVar()
French = Checkbutton(window, text='French', var=French_state)
French.grid(column=1, row=14)
Frenchcombo = Combobox(window)
Frenchcombo['values']= (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
Frenchcombo.grid(column=2, row=14)
Frenchcombo.current(19)
Average_of_French1 = Label(window, text="Average_of_French: ")
Average_of_French1.grid(column=4, row=14)
Average_of_French2 = Label(window, text="-")
Average_of_French2.grid(column=5, row=14)







secheader = Label(window, text="Grade : ")
secheader.grid(column=0, row=15)

v = IntVar()
rad1 = Radiobutton(window,text='1st',variable=v, value=1)
rad2 = Radiobutton(window,text='2nd',variable=v, value=2)
rad3 = Radiobutton(window,text='3rd',variable=v, value=3)
rad4 = Radiobutton(window,text='4th',variable=v, value=4)
rad5 = Radiobutton(window,text='5th',variable=v, value=5)
rad1.grid(column=0, row=16)
rad2.grid(column=1, row=16)
rad3.grid(column=2, row=16)
rad4.grid(column=3, row=16)
rad5.grid(column=4, row=16)




txt = scrolledtext.ScrolledText(window,width=120,height=25)
txt.place(x=12, y=360)







def clicked():
    global sdate
    sdate = selecteddate
    global ready_to_add
    ready_to_add = True
    txt.delete('1.0', END)

    if stnametxt.get():
        pass
    else:
        messagebox.askretrycancel('You have to fill all of the items !','Please enter First Name')
        ready_to_add = False


    if (stlnametxt.get()):
        pass
    else:
        messagebox.askretrycancel('You have to fill all of the items !','Please enter Last Name')
        ready_to_add = False
 
 
    if sdate != 1:
        pass
    else:
        messagebox.askretrycancel('You have to fill all of the items !','Please select the Birthday')
        ready_to_add = False
        
    
    if v.get():
        pass
    else:
        messagebox.askretrycancel('You have to fill all of the items !','Please select the Grade !')
        ready_to_add = False
    






    global xxx
    xxx = ready_to_add
    
    if (xxx):
        courselist = []

        if (Art_state.get()):
            courselist.append(Artcombo.get())
        else:
            courselist.append(False)

        if Mathematics_state.get() :
            courselist.append(Mathematicscombo.get())
        else:
            courselist.append(False)        
        
        if Music_state.get() :
            courselist.append(Musiccombo.get())
        else:
            courselist.append(False)
        
        if Dance_state.get() :
            courselist.append(Dancecombo.get())
        else:
            courselist.append(False)
        
        if PhysicalScience_state.get() :
            courselist.append(PhysicalSciencecombo.get())
        else:
            courselist.append(False)
        
        if EnglishLiterature_state.get() :
            courselist.append(EnglishLiteraturecombo.get())
        else:
            courselist.append(False)
        
        if Chemistry_state.get() :
            courselist.append(Chemistrycombo.get())
        else:
            courselist.append(False)
    
        if French_state.get() :
            courselist.append(Frenchcombo.get())
        else:
            courselist.append(False)
 














    
        x = datetime.datetime.now()
        newstid = x.strftime("%Y%m%d%S%f")
        allstudents[newstid] = (stnametxt.get(), stlnametxt.get(),sdate, v.get(), courselist)
        
        for i in allstudents:
            how_many_courses = 0
            total_of_notes = 0
            
            txt.insert(INSERT,'\nStudent ID:   ')
            txt.insert(INSERT,i)
            txt.insert(INSERT,'\n\nFirst name: ')
            txt.insert(INSERT,allstudents[i][0])
            txt.insert(INSERT,'     Last name: ')
            txt.insert(INSERT,allstudents[i][1])
            txt.insert(INSERT,'\n\n  Birthday: ')
            txt.insert(INSERT,allstudents[i][2])
            txt.insert(INSERT,'\n  Grade: ')
            txt.insert(INSERT,allstudents[i][3])
            txt.insert(INSERT,'\n\n  Scores: ')
            
            if allstudents[i][4][0]:
                total_of_notes = total_of_notes + int(allstudents[i][4][0])
                how_many_courses +=1
                txt.insert(INSERT,'\n     Art: ')
                txt.insert(INSERT,allstudents[i][4][0])
            
            if allstudents[i][4][1]:
                total_of_notes = total_of_notes + int(allstudents[i][4][1])
                how_many_courses +=1
                txt.insert(INSERT,'\n     Mathematics: ')
                txt.insert(INSERT,allstudents[i][4][1])
            
            if allstudents[i][4][2]:
                total_of_notes = total_of_notes + int(allstudents[i][4][2])
                how_many_courses +=1
                txt.insert(INSERT,'\n     Music: ')
                txt.insert(INSERT,allstudents[i][4][2])
            
            if allstudents[i][4][3]:
                total_of_notes = total_of_notes + int(allstudents[i][4][3])
                how_many_courses +=1
                txt.insert(INSERT,'\n     Dance: ')
                txt.insert(INSERT,allstudents[i][4][3])
            
            if allstudents[i][4][4]:
                total_of_notes = total_of_notes + int(allstudents[i][4][4])
                how_many_courses +=1
                txt.insert(INSERT,'\n     Physical Science: ')
                txt.insert(INSERT,allstudents[i][4][4])
            
            if allstudents[i][4][5]:
                total_of_notes = total_of_notes + int(allstudents[i][4][5])
                how_many_courses +=1
                txt.insert(INSERT,'\n     EnglishLiterature: ')
                txt.insert(INSERT,allstudents[i][4][5])
            
            if allstudents[i][4][6]:
                total_of_notes = total_of_notes + int(allstudents[i][4][6])
                how_many_courses +=1
                txt.insert(INSERT,'\n     Chemistry: ')
                txt.insert(INSERT,allstudents[i][4][6])
            
            if allstudents[i][4][7]:
                total_of_notes = total_of_notes + int(allstudents[i][4][7])
                how_many_courses +=1
                txt.insert(INSERT,'\n     French: ')
                txt.insert(INSERT,allstudents[i][4][7])
            txt.insert(INSERT,'\n\n          Total of the selected courses: ')
            txt.insert(INSERT,how_many_courses)
            txt.insert(INSERT,'\n\n          Average: ')
            txt.insert(INSERT,(total_of_notes/how_many_courses))
            txt.insert(INSERT,'\n________________________________________________________________________________________________________________________\n')

    
            
    global total
    global counttt
    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][0]:
            total = total + int(allstudents[i][4][0])
            counttt += 1
            Average_of_Art2.configure(text= (total/counttt))
        
       

    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][1]:
            total = total + int(allstudents[i][4][1])
            counttt += 1
            Average_of_Mathematics2.configure(text= (total/counttt))
        
       
        

         

    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][2]:
            total = total + int(allstudents[i][4][2])
            counttt += 1
            Average_of_Music2.configure(text= (total/counttt))
        
       
        

         

    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][3]:
            total = total + int(allstudents[i][4][3])
            counttt += 1
            Average_of_Dance2.configure(text= (total/counttt))
        
       
        

         

    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][4]:
            total = total + int(allstudents[i][4][4])
            counttt += 1
            Average_of_PhysicalScience2.configure(text= (total/counttt))
        
       
        

         

    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][5]:
            total = total + int(allstudents[i][4][5])
            counttt += 1
            Average_of_EnglishLiterature2.configure(text= (total/counttt))
        
       
        

         

    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][6]:
            total = total + int(allstudents[i][4][6])
            counttt += 1
            Average_of_Chemistry2.configure(text= (total/counttt))
        
       
        
       

    total = 0
    counttt = 0
    for i in allstudents:
        if allstudents[i][4][7]:
            total = total + int(allstudents[i][4][7])
            counttt += 1
            Average_of_French2.configure(text= (total/counttt))
        
       
        

  
        
    
    
    #Empty the input boxes after add a student
    #stnametxt.configure(text= ' ')
    #stlnametxt.configure(text= ' ')
    #studentidtxt.configure(text= ' ')




btn = Button(window, text="Add", command=clicked)
btn.grid(column=5, row=16)

 
 
 
 
 
 
 
 
 

 
 
 
 
 
 
 
 
 

window.mainloop()