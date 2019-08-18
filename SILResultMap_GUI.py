from tkinter import *
import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter import messagebox
from tkinter import ttk
import CsvResultChecker
class GUI_COntroller:
    '''
	   This class initialize the required controls for TkInter GUI
	'''
    def __init__(self,TkObject):
 
 
	    #Load company image
        Imageloc=tk.PhotoImage(file='logo_gif.gif')		
        label3=Label(image=Imageloc,)
        label3.image = Imageloc		
        label3.place(x=200,y=30)
		
        #label
        label_MyName = Label(TkObject,bd=7, text="For any clarification on this tool contact", bg='gray', fg="black",font=200)	
        label_MyName.place(x=20,y=450)

        #label
        label_MyName = Label(TkObject,bd=7, text="masthanvali.s@silver-atena.com", bg='gray', fg="blue",font=200)
        label_MyName.place(x=300,y=450)

		
        global TkObject_ref
        TkObject_ref =  TkObject       
		
        #label
        global label1		
        label1 = Label(TkObject,bd=7, text="Browse Path:", bg="yellow", fg="black",font=200)	
        label1.place(x=50,y=130)

        #select file
        global 	button1	
        button1=Button(TkObject,activebackground='green',borderwidth=5, text='Browse SIL results path',width=20, command=GUI_COntroller.openfile)
        button1.place(x=230,y=130)

        #Exit Window
        global button2		
        button2=Button(TkObject,activebackground='green',borderwidth=5, text='Close Window', command=GUI_COntroller.exitWindow)
        button2.place(x=450,y=130)	

    def exitWindow():
        	 TkObject_ref.destroy()

    def openfile():
        global filepath
        filepath = askdirectory()

        label6= Label(TkObject_ref,bg='orange',text='Selected Path: ',font=40)
        label6.place(x=30,y=250)		
		
        EntryObj2 = Entry(TkObject_ref,font=10)
        EntryObj2.place(x=170,y=250,width=350)
        EntryObj2.insert(0,filepath)
        EntryObj2.configure(state='readonly')		
		
        button5 = Button(TkObject_ref,text='Run Test',font=10,bd=5,command=TestScript.RunTest)
        button5.place(x=230,y=320)
			
	
class TestScript:
    def RunTest():

        CsvResultChecker.ResultChecker(filepath,TkObject_ref)
		

def SILResultMap():
	
       root = tk.Tk()
       
       #Change the background window color
       root.configure(background='gray')     
       
       #Set window parameters
       root.geometry('550x500')
       root.title('Welcome to SIL Results check tool')
       
       #Removes the maximizing option
       root.resizable(0,0)
       
       ObjController = GUI_COntroller(root)
       
       #keep the main window is running
       root.mainloop()
