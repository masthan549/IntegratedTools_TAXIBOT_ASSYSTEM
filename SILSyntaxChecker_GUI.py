from tkinter import *
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from tkinter import ttk
import SIL_ScriptSyntaxChecker
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
        label1.place(x=50,y=150)

        #select file
        global 	button1	
        button1=Button(TkObject,activebackground='green',borderwidth=5, text='Select script file!!',width=15, command=GUI_COntroller.openfile)
        button1.place(x=230,y=150)

        #Exit Window
        global button2		
        button2=Button(TkObject,activebackground='green',borderwidth=5, text='Close Window', command=GUI_COntroller.exitWindow)
        button2.place(x=450,y=150)	

    def exitWindow():
        	 TkObject_ref.destroy()

    def openfile():
        global filepath,filepath_temp	
        filepath = askopenfilename()
        filepath_temp=filepath.split('/')
        filepath_temp=filepath_temp[len(filepath_temp)-1]
		
        print(filepath_temp.split('.')[len(filepath_temp.split('.'))-1:][0])
        if not filepath_temp.endswith('txt'):
            messagebox.showerror('Error','Select only txt file!!')
            TkObject_ref.destroy()			
		
        if len(filepath):
            label1.destroy()
            button1.destroy()			
            label4= Label(TkObject_ref,bg='orange',text='Selected SIL script file: '+filepath_temp,font=40)
            label4.place(x=20,y=100)

            button5 = Button(TkObject_ref,text='Run Test',font=10,bd=5,command=TestScript.RunTest)
            button5.place(x=330,y=150)
			
            global label5
            label5= Label(TkObject_ref,bg='orange',text='Enter System execution time: ',font=40)
            label5.place(x=30,y=200)

            global EntryObj
            EntryObj = Entry(TkObject_ref,font=10,bd=5)
            EntryObj.place(x=330,y=200)
			
            global label6
            label6= Label(TkObject_ref,bg='orange',text='Enter SIL global request number: ',font=40)
            label6.place(x=30,y=270)

            global EntryObj2
            EntryObj2 = Entry(TkObject_ref,font=10,bd=5)
            EntryObj2.place(x=330,y=270)			
	
class TestScript:
    def RunTest():
	
        exeTime = EntryObj.get()
        SILreqno = EntryObj2.get()

        if len(exeTime):
            if len(SILreqno):
                SIL_ScriptSyntaxChecker.script_exe(filepath,filepath_temp,TkObject_ref,exeTime,SILreqno)
            else:
                messagebox.showerror('Error','Enter SIL global request number')
        else:
            messagebox.showerror('Error','Please Enter System execution time')
		

def SILSyntaxChecker():
	
       root = tk.Tk()
       
       #Change the background window color
       root.configure(background='gray')     
       
       #Set window parameters
       root.geometry('550x500')
       root.title('Welcome to SIL Syntax checker Tool')
       
       #Removes the maximizing option
       root.resizable(0,0)
       
       ObjController = GUI_COntroller(root)
       
       #keep the main window is running
       root.mainloop()
