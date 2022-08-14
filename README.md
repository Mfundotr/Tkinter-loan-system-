# Tkinter-loan-system-
from tkinter import *
from tkcalendar import * 
from openpyxl import workbook, load_workbook
from tkinter import messagebox


def write():
    firstname_info=firstname.get()
    lastname_info=lastname.get()
    surname_info=surname.get()
    gender1_info=  text=checkvar1.get()
    gender2_info= text=checkvar2.get()
    kind= str(gender1_info) +str(gender2_info)
    age_info=age1.get()
    id_info=id_entry.get()
    phonenumber_info=phonenumber.get()
    addressnumber_info=address_number.get()
    creditrequest_info=credit_request.get()
    payday_info=pay_day.get_date()
    total=40/1*int(creditrequest_info) /100 +int(creditrequest_info)
   
    
    file = open("loaners.txt", "a")
    file.write("\n _______________________") 
    file.write("\n First_name: " + firstname_info)
    file.write("\n Last name: " + lastname_info) 
    file.write("\n Surname: "+ surname_info)
    file.write("\n Gender: "+kind) 
    file.write("\n Age: "+age_info)
    file.write("\n ID_Number: " +id_info)
    file.write("\n Phone_number: " +phonenumber_info)
    file.write("\n Address_number: " + addressnumber_info)
    file.write("\n Credit_request: " + creditrequest_info) 
    file.write("\n With Interest: " +str(total)) 
    file.write("\n Pay_Day: " +payday_info) 
    file.close()
    wb= load_workbook('loaners.xlsx')
    ws=wb.active
    ws['A1']="First_N"
    ws['B1'] ="Last_N"
    ws['C1'] ="Surname"
    ws['D1'] ="Gender"
    ws['E1'] ="Age"
    ws['F1'] ="ID_Num"
    ws['G1']="Phone_N"
    ws['H1']="Address"
    ws['I1']="Credit"
    ws['J1']="interest"
    ws['K1']="PayDay"
   
    ws.append([firstname_info,lastname_info,surname_info,kind , age_info , id_info, phonenumber_info, addressnumber_info,creditrequest_info,total , payday_info])
    wb.save('loaners.xlsx')
    
    firstname_entry.delete(0, END)
    lastname_entry.delete(0, END)
    surname_entry.delete(0, END)
    address_entry.delete(0 ,END)
    phonenumber.delete(0, END)
    age_entry.delete(0,END) 
    credit_entry.delete(0, END)
    id_entry.delete(0,END)
  


root=Tk()
root.title("Loan app")
root.geometry("500x400")
root.maxsize(1000,800 )

    

intro=Label(root, text="Loan Application Program:",  bg="grey",fg="black", width="500", height="3") 
intro.pack()

first_name= Label(root,text="Enter First Name:")
first_name.place(x=10, y=175)
last_name= Label (root, text="Enter Last Name:")
last_name.place(x=10, y=245)
surname =Label (root, text ="Enter Surname:")
surname.place(x=10,y =317)
phone_number=Label(root, text="Phone Number :") 
phone_number.place(x=10, y=648)

age=Label(root, text="Enter Age :") 
age.place(x=10, y=450)
id=Label(root, text="Enter ID Number:")
id.place(x=10, y=493)
address_number=Label(root, text="Address Number :") 

address_number.place(x=10, y=570)
pd=Label(root,text="Pay Day")
pd.place(x=515,y=280)

credit_request=Label(root, text="Enter Credit_request:")
credit_request.place(x=515,y=175)
pay_day= Calendar(root,selectmode="day", year=2022, month=6, day =8)
pay_day.place(x=515, y=300)
text=Label(root, text ="" ,  bg="grey",fg="black", width="500", height="1")  
text.place(x=1, y=850)

firstname =StringVar()
lastname=StringVar()
surname=StringVar()
age=IntVar()
phonenumber=StringVar()
address_number=StringVar ()
credit_request=StringVar()
checkvar1 = StringVar()
checkvar2 = StringVar()
age1=StringVar()

firstname_entry=Entry(textvariable =firstname)
firstname_entry.place(x=10,y=210, width=440) 
lastname_entry=Entry(textvariable =lastname)
lastname_entry.place(x=10,y =280, width=440)
surname_entry=Entry(textvariable =surname )
surname_entry.place(x=10,y=355, width=440) 
c1= Checkbutton(text = "Male", variable = checkvar1,onvalue = "Male" , offvalue = "" ,) 
c1.deselect()
c1.place(x=130,y=400)
c2= Checkbutton(text = "Female", variable = checkvar2,onvalue = "Female" , offvalue = "" ,)
c2.deselect()
c2.place(x=190,y=400)
age_entry=Entry(textvariable =age1)
age_entry.place(x=130,y=450, width=40) 
id_entry=Entry(textvariable=id)
id_entry.place(x="10",y="533", width =440)
  
address_entry=Entry(textvariable=address_number) 

address_entry.place(x=10,y=610, width=440)
phonenumber= Entry(root, textvariable="phone_number")
phonenumber.place(x=10,y=680, width=440)
credit_entry=Entry (textvariable =credit_request)
credit_entry.place(x=515, y=210, width=300,) 
button=Button(text ="Submit", command=write)
button.place(x=600,y=550)

def info():
   messagebox.showinfo("Loan Entry Application","Please Fill In Acordingly")


b= Button(root, text ="info", relief=RAISED, command=info, bitmap="info")
b.place(x=100, y=135)



footer=Label(root, text="",  bg="grey",fg="black", width="500", height="2") 
footer.place(x=0,y=1390)
root.mainloop()
