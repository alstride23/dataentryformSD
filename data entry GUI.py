#This program is designed to open an excel spreadsheet and 
#write to the spreadsheet character information for the user.
#The user is able to input specific information about who they
#are needing to input and then confirming the entries.

from tkinter import *
import openpyxl


infoSheet = Tk(className= 'Character Entry Sheet')

infoSheet.geometry ("700x500")

infoSheet.config(background = 'grey')

label = Label(infoSheet, font='calibri', text="Enter Character Name: ")
label.pack()

entry = Entry(infoSheet, font='calibri', fontsize='14', justify="center")
entry.pack()

infoSheet.mainloop()