from tkinter import *
import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter import messagebox
from tkinter import ttk

import TP2RTRTTS_GUI
import DoorsTPToTemplate_GUI
import TP2SILTS_GUI
import SILResultMap_GUI
import SILSyntaxChecker_GUI
import DoorsTCToTemplate_GUI
import TCToDoors_GUI
import TPToDoors_GUI

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

        #label
        label_MyName2 = Label(TkObject,bd=7, text="Select Tool which you want to use: ", bg='white', fg="purple",font=400)	
        label_MyName2.place(x=20,y=130)
		
		

        chkObj6 = Checkbutton(TkObject, text='1. Doors TC to xlsx Template', command=GUI_COntroller.DoorsTCToTemplate, bg='Green')
        chkObj6.place(x=50,y=200)	
		
        chkObj7 = Checkbutton(TkObject, text='2. TC Template to Doors', command=GUI_COntroller.TCToDoorsTemplate, bg='Green')
        chkObj7.place(x=50,y=250)	
		
        chkObj4 = Checkbutton(TkObject, text='3. Doors TP to xlsx Template', command=GUI_COntroller.DoorsToTemplate, bg='Yellow')
        chkObj4.place(x=50,y=320)	

        chkObj8 = Checkbutton(TkObject, text='4. TP Template to Doors', command=GUI_COntroller.TPToDoorsTemplate, bg='Yellow')
        chkObj8.place(x=50,y=370)			
		
		
		
		
        chkObj1 = Checkbutton(TkObject, text='5. TP To RTRT script Generator', command=GUI_COntroller.TPToRTRT, bg='Orange')
        chkObj1.place(x=350,y=200)
		
        chkObj2 = Checkbutton(TkObject, text='6. TP To SIL script Generator', command=GUI_COntroller.TPToSIL, bg='Orange')
        chkObj2.place(x=350,y=250)			

        chkObj3 = Checkbutton(TkObject, text='7. SIL Results map Tool', command=GUI_COntroller.TPToSILMap, bg='Light Blue')
        chkObj3.place(x=350,y=320)		

        chkObj5 = Checkbutton(TkObject, text='8. SIL Script Syntax Checker', command=GUI_COntroller.SILScriptSyntaxChecker, bg='Light Blue')
        chkObj5.place(x=350,y=370)	

        global TkObject_ref
        TkObject_ref =  TkObject
		
        global button2		
        button2=Button(TkObject,activebackground='green',borderwidth=5, text='Close Window', command=GUI_COntroller.exitWindow)
        button2.place(x=450,y=130)	

    def exitWindow():
        	 TkObject_ref.destroy()

    def TPToRTRT():
        TkObject_ref.destroy()	
        TP2RTRTTS_GUI.TPToRTRT()
		
    def TPToSIL():       
        TkObject_ref.destroy()
        TP2SILTS_GUI.TP2SILTS()

    def TPToSILMap():       
        TkObject_ref.destroy()
        SILResultMap_GUI.SILResultMap()

    def DoorsToTemplate():       
        TkObject_ref.destroy()
        DoorsTPToTemplate_GUI.DoorsTPToTemplate()

    def SILScriptSyntaxChecker():
        TkObject_ref.destroy()
        SILSyntaxChecker_GUI.SILSyntaxChecker()
		
    def DoorsTCToTemplate():
        TkObject_ref.destroy()
        DoorsTCToTemplate_GUI.DoorsTCToTemplateConv()
		
    def TCToDoorsTemplate():
        TkObject_ref.destroy()
        TCToDoors_GUI.TCTemplateToDoors()

    def TPToDoorsTemplate():
        TkObject_ref.destroy()
        TPToDoors_GUI.TPTemplateToDoors()
		
if __name__ == '__main__':	
	
       root = tk.Tk()
       
       #Change the background window color
       root.configure(background='gray')     
       
       #Set window parameters
       root.geometry('550x500')
       root.title('Welcome Taxibot Integrated Tool environment')
       
       #Removes the maximizing option
       root.resizable(0,0)
       
       ObjController = GUI_COntroller(root)
       
       #keep the main window is running
       root.mainloop()
