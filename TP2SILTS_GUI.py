from tkinter import *
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from tkinter import ttk
import TP2SIL_TS
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
        label_MyName.place(x=20,y=500)

        #label
        label_MyName = Label(TkObject,bd=7, text="masthanvali.s@silver-atena.com", bg='gray', fg="blue",font=200)
        label_MyName.place(x=300,y=500)
		
	    #SAEntry = Entry(root,takefocus=False,justify=tk.CENTER,font=50,)

        global TkObject_ref
        TkObject_ref =  TkObject       
		
        #label
        global label1		
        label1 = Label(TkObject,bd=7, text="Select the xls or xlsx file which has Test Script:", bg="yellow", fg="black",font=200)	
        label1.place(x=50,y=130)

        #select file
        global 	button1	
        button1=Button(TkObject,activebackground='green',borderwidth=5, text='Select file!!',width=10, command=GUI_COntroller.openfile)
        button1.place(x=430,y=130)

        #Exit Window
        global button2		
        button2=Button(TkObject,activebackground='green',borderwidth=5, text='Close Window', command=GUI_COntroller.exitWindow)
        button2.place(x=550,y=130)	

    def exitWindow():
        	 TkObject_ref.destroy()

    def openfile():
        global filepath,filepath_temp	
        filepath = askopenfilename()
        filepath_temp=filepath.split('/')
        filepath_temp=filepath_temp[len(filepath_temp)-1]
		
        if not (filepath_temp.endswith('xls') or filepath_temp.endswith('xlsx') or filepath_temp.endswith('xlsm')):
            messagebox.showerror('Error','Select only xls/xlsx file!!')
            TkObject_ref.destroy()			
		
        if len(filepath):
            label1.destroy()
            button1.destroy()			
            label4= Label(TkObject_ref,bg='orange',text='Selected file: '+filepath_temp,font=40)
            label4.place(x=30,y=130)

            button5 = Button(TkObject_ref,text='Run Test',font=10,bd=5,command=TestScript.RunTest)
            button5.place(x=430,y=130)
			
            global label5
            label5= Label(TkObject_ref,bg='orange',text='Enter Test script sheet Index in selected file: ',font=40)
            label5.place(x=30,y=200)

            global EntryObj
            EntryObj = Entry(TkObject_ref,font=10,bd=5)
            EntryObj.place(x=530,y=200)
			
            global label6
            label6= Label(TkObject_ref,bg='orange',text='Enter your system execution time: ',font=40)
            label6.place(x=30,y=270)

            global EntryObj2
            EntryObj2 = Entry(TkObject_ref,font=10,bd=5)
            EntryObj2.place(x=530,y=270)	

            #global request number
            global label61
            label61= Label(TkObject_ref,bg='orange',text='Enter SIL global request number: ',font=40)
            label61.place(x=30,y=350)

            global EntryObj23
            EntryObj23 = Entry(TkObject_ref,font=10,bd=5)
            EntryObj23.place(x=530,y=350)

            #global request number
            global label62
            label62= Label(TkObject_ref,bg='orange',text='Enter SIL script execution start time: ',font=40)
            label62.place(x=30,y=420)

            global EntryObj22
            EntryObj22 = Entry(TkObject_ref,font=10,bd=5)
            EntryObj22.place(x=530,y=420)	

            # Validate the given inputs for execution
            			
	
class TestScript:
    def RunTest():
	
        dataGiven = EntryObj.get()
        systemExeTime = EntryObj2.get()
        SILGlobNumber = EntryObj23.get()
        SILScriptStartTime = EntryObj22.get()

        if len(dataGiven):
            if len(systemExeTime):
                if len(SILGlobNumber):
                    if len(SILScriptStartTime):
                        TP2SIL_TS.GenerateTS(filepath,filepath_temp,TkObject_ref,dataGiven,systemExeTime,SILGlobNumber,SILScriptStartTime)
                    else:
                        messagebox.showerror('Error','Please Enter SIL script execution start time')
                else:
                    messagebox.showerror('Error','Please enter SIL global request number')
            else:
                messagebox.showerror('Error','Please enter your system execution time')
        else:
            messagebox.showerror('Error','Please enter Test script sheet Index in selected file')
		

def TP2SILTS():	
	
    root = tk.Tk()
    
    #Change the background window color
    root.configure(background='gray')     
    
    #Set window parameters
    root.geometry('850x600')
    root.title('Welcome to SIL script generator tool')
    
    #Removes the maximizing option
    root.resizable(0,0)
    
    ObjController = GUI_COntroller(root)
    
    #keep the main window is running
    root.mainloop()
