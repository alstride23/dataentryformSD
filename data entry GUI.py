#This program is designed to open an excel spreadsheet and 
#write to the spreadsheet character information for the user.
#The user is able to input specific information about who they
#are needing to input and then confirming the entries.

from tkinter import *
import openpyxl
from tkinter import ttk

def enter_data():
    firstname= first_name_entry.get()
    lastname= last_name_entry.get()
    pronouns = pronouns_dropbox.get()
    age= age_spinbox.get()
    race= race_dropbox.get()
    classes = classes_dropbox.get()
    
    print("First name: ", firstname, "Last name: ", lastname)
    print("Pronouns:", pronouns , "Race: ", race ,"Class: ", classes )


infoSheet = Tk(className= 'Character Entry Sheet')

#Character Information


infoSheet_frame = LabelFrame(infoSheet, text="Character Information")
infoSheet_frame.grid(row=0, column=0, padx=20, pady=20)


first_name_label = Label(infoSheet_frame, text = "First Name")
first_name_label.grid(row=0, column=0)
last_name_label= Label(infoSheet_frame, text = "Last Name")
last_name_label.grid(row=0,column=1)

first_name_entry = Entry(infoSheet_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry= Entry(infoSheet_frame)
last_name_entry.grid(row=1, column=1)

pronouns_label= Label(infoSheet_frame, text = "Pronouns")
pronouns_dropbox = ttk.Combobox(infoSheet_frame, values= ["She/Her", "He/Him", "They/Them","Other"])
pronouns_label.grid(row =0, column = 2)
pronouns_dropbox.grid(row=1,column=2)

age_label = Label(infoSheet_frame, text="Age")
age_spinbox = Spinbox(infoSheet_frame, from_=0, to="infinity")
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

race_label = Label(infoSheet_frame, text="Race")
race_dropbox = ttk.Combobox(infoSheet_frame, values=["Human", "Elf", "Half-Elf", "Tiefling", "Dwarf", "Orc", "Dragonborn", "Gnome", "Other"])
race_label.grid(row=2, column= 1)
race_dropbox.grid(row=3, column=1)

classes_label = Label(infoSheet_frame, text="Class")
classes_dropbox= ttk.Combobox(infoSheet_frame, values=["Artificier", "Barbarian", "Bard", "Cleric", "Druid", "Fighter", "Monk", "Paladin", "Ranger", "Rogue", "Sorceror", "Warlock", "Wizard"])
classes_label.grid(row=2, column=2)
classes_dropbox.grid(row=3,column=2)

for widget in infoSheet_frame.winfo_children():
    widget.grid_configure(padx=5, pady=5)

#Stats

statistics_frame=LabelFrame(infoSheet, text="Character Stats")
statistics_frame.grid(row=5, column=0, padx=20, pady=20)

strength_label=Label(statistics_frame, text="Strength")
strength_label.grid(row=2, column=0)
strength_spinbox= Spinbox(statistics_frame, from_=0, to="Infinity")
strength_spinbox.grid(row=2, column=1)

dexterity_label=Label(statistics_frame, text="Dexterity")
dexterity_label.grid(row=3, column=0)
dexterity_spinbox=Spinbox(statistics_frame, from_=0, to="infinity")
dexterity_spinbox.grid(row=3, column=1)

constitution_label=Label(statistics_frame, text="Constitution")
constitution_label.grid(row=4, column=0)
constitution_spinbox=Spinbox(statistics_frame, from_=0, to="infinity")
constitution_spinbox.grid(row=4, column=1)

intelligence_label=Label(statistics_frame, text= "Intelligence")
intelligence_label.grid(row=5, column=0)
intelligence_spinbox=Spinbox(statistics_frame, from_=0, to="infinity")
intelligence_spinbox.grid(row=5, column=1)

wisdom_label=Label(statistics_frame, text="Wisdom")
wisdom_label.grid(row=6, column=0)
wisdom_spinbox=Spinbox(statistics_frame, from_=0,to="infinity")
wisdom_spinbox.grid(row=6, column=1)

charisma_label=Label(statistics_frame, text="Charisma")
charisma_label.grid(row=7, column=0)
charisma_spinbox=Spinbox(statistics_frame, from_=0, to="infinity")
charisma_spinbox.grid(row=7, column=1)

for widget in statistics_frame.winfo_children():
    widget.grid_configure(padx=10, pady=10)



infoSheet.mainloop()
