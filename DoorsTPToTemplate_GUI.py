from tkinter import *
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from tkinter import ttk
import DoorsTPToTemplate_Script
from threading import *

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

		
	    #SAEntry = Entry(root,takefocus=False,justify=tk.CENTER,font=50,)

        global TkObject_ref
        TkObject_ref =  TkObject       
		
        #label
        global label1		
        label1 = Label(TkObject,bd=7, text="Select the file:", bg="yellow", fg="black",font=200)
        label1.place(x=50,y=130)

        #select file
        global 	button1	
        button1=Button(TkObject,activebackground='green',borderwidth=5, text='Select file!!',width=10, command=GUI_COntroller.openfile)
        button1.place(x=230,y=130)

        #Exit Window
        global button2		
        button2=Button(TkObject,activebackground='green',borderwidth=5, text='Close Window', command=GUI_COntroller.exitWindow)
        button2.place(x=450,y=130)	

    def exitWindow():
        	 TkObject_ref.destroy()

    def openfile():
        global filepath,filepath_temp	
        filepath = askopenfilename()
        filepath_temp=filepath.split('/')
        filepath_temp=filepath_temp[len(filepath_temp)-1]
		
        if not (filepath_temp.endswith('xls') or filepath_temp.endswith('xlsx')):
            messagebox.showerror('Error','Select only xls/xlsx file!!')
            TkObject_ref.destroy()			
		
        if len(filepath):
            label1.destroy()
            button1.destroy()
            label4= Label(TkObject_ref,bg='orange',text='Selected file: '+filepath_temp,font=40)
            label4.place(x=30,y=130)

            button5 = Button(TkObject_ref,text='Convert To Template',font=10,bd=5,command=TestScript.RunTest)
            button5.place(x=250,y=200)

			
class TestScript:
    def RunTest():
        DoorsTPToTemplate_Script.startFun(filepath,filepath_temp,TkObject_ref)

def DoorsTPToTemplate():
	
       root = tk.Tk()
       
       #Change the background window color
       root.configure(background='gray')     
       
       #Set window parameters
       root.geometry('550x500')
       root.title('Welcome to Doors TP to xlsx Template conversion Tool')
       
       #Removes the maximizing option
       root.resizable(0,0)
	   
       ObjController = GUI_COntroller(root)
       
       #keep the main window is running
       root.mainloop()       
