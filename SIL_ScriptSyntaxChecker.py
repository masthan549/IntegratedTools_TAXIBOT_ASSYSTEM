import sys,os,re,copy
from tkinter import *

errorsLog = []
lineNumbers = []

def OpenFile():
    try:
        global ScriptFile
        ScriptFile = open(fileName)
    except:
        messagebox.showerror('Error','Error occured while opening selected file!!')
        GUIObj.destroy()
        sys.exit()


def fileHeaderCheck():
    '''
	   Script file header should be started at 1 line and ended at 10th line(with as-trick)
	'''


    HeaderStartingIndex = 1
    HeaderEndingIndex = 10
    count = 1
	
    try:
        for lineInfo in ScriptFile:
            if count == 1:
                if len(str(re.search('\*.+\*',lineInfo))) == 4:
                    errorsLog.append('Script header should be started in first line')
                    lineNumbers.append('1')				
            if count == 10:
                if len(str(re.search('[\*.+\*]',lineInfo))) <= 8:
                    errorsLog.append('Script header should be ended at 10th line')				
                    lineNumbers.append('10')				
                break				
            count = count+1
    except:
        messagebox.showerror('Error','Script file header should be started at 1 line and ended at 10th line!!')
        GUIObj.destroy()
        sys.exit()

def SignalDescription():
    '''
        1) After header signal naming format "  step,   time,  type,  num_of_signals, signal_names" should be there.
		2) Next line format should be "STEP#,  TIME,  TYPE,  <Number of signals>, GLOB_REQ_BUS7:SilGlobReq, GLOB_REQ_BUS7:SilGlobCounter, <input_Signals>"
    '''	
    count = 1
    ScriptFile.seek(0,0)
    signalDes = ['step', 'time', 'type', 'num_of_signals', 'signal_names']
    for lineInfo in ScriptFile:
        if count == 11:
            if len(lineInfo)<=40:
                messagebox.showerror('Error','After immediate to the Script file header, must start "step,   time,  type,  num_of_signals, signal_names"')
                GUIObj.destroy()
                sys.exit()
            else:
                AfterSplit = lineInfo.split(',')
                if len(AfterSplit) == 5:
                    for indx,data in enumerate(AfterSplit):
                        if not signalDes[indx] in data:
                            errorsLog.append('Signal name "'+data+'" does not match with the signal name in "step,   time,  type,  num_of_signals, signal_names"')
                            lineNumbers.append('11')
                else:

                    errorsLog.append('Signal names must be in the format of "step,   time,  type,  num_of_signals, signal_names"')
                    lineNumbers.append('11')
        if count == 12:
            if len(lineInfo) <= 80:
                messagebox.showerror('Error','At 12th line should have "STEP#,  TIME,  TYPE,  <Number of signals>, GLOB_REQ_BUS7:SilGlobReq, GLOB_REQ_BUS7:SilGlobCounter, <input_Signals>"')
                GUIObj.destroy()
                sys.exit()
            else:
                AfterSplit = lineInfo.split(',')
                try: 				
                    if 'STEP#' in AfterSplit[0] and 'TIME' in AfterSplit[1] and 'TYPE' in AfterSplit[2] and 'GLOB_REQ_BUS7:SilGlobReq' in  AfterSplit[4] and 'GLOB_REQ_BUS7:SilGlobCounter' in  AfterSplit[5]:
                        AfterSplit.pop()
                        global noOfSignals						
                        noOfSignals = int(AfterSplit[3])						
                        if not len(AfterSplit[4:]) == noOfSignals:
                            errorsLog.append('Number of signal does not match with the number given to ''num_of_signals')
                            lineNumbers.append('12')
                    else:					
                        errorsLog.append('At 12th line should have "STEP#,  TIME,  TYPE,  <Number of signals>, GLOB_REQ_BUS7:SilGlobReq, GLOB_REQ_BUS7:SilGlobCounter, <input_Signals>"')
                        lineNumbers.append('12')
                except:
                        messagebox.showerror('Error','At 12th line should have "STEP#,  TIME,  TYPE,  <Number of signals>, GLOB_REQ_BUS7:SilGlobReq, GLOB_REQ_BUS7:SilGlobCounter, <input_Signals>"')
                        GUIObj.destroy()
                        sys.exit()
                     				
            break
        count = count+1
		
def SyntaxCheckAfterPowerOnDelay():
    '''
        1) Check the number of signals are matching with the given number
		2) Check the spaces(should not have tabs)
		3) Only one space after every comma(,) except for the step after timer in every cycle.
		4) time should be sequential and Should only start with C and end with D.
    '''	
    isNoScript = True	
    isScriptHasD = False
    Dstarted = False
    counter = 1	
    PrevCycleTime = 0
    cycleDelay = False	 
    endOfexec = False
    global prevCounterValue	
    prevCounterValue = 0
    try:	
        for lineIndx in ScriptFile:
            endOfexec = False	
            counter = counter+1	
            if isScriptHasD is True:
                Dstarted = False
            lineDataAfterParse = lineIndx.split(',')
            try:
                if lineDataAfterParse[len(lineDataAfterParse)-1:][0] == '' or (re.search('[\s]+',lineDataAfterParse)):
                    lineDataAfterParse.pop()
            except: pass
        
            if len(lineDataAfterParse) >=4:
                isNoScript = False
                if ('D' in lineDataAfterParse[2]):
                    Dstarted = True			
                    isScriptHasD = True
                    if not (noOfSignals == (len(lineDataAfterParse)-4)):
                        errorsLog.append('Number signals should be '+str(noOfSignals)+', given number of signals are:'+str(len(lineDataAfterParse)-4))
                        lineNumbers.append(str(counter+11))
                elif ('C' in lineDataAfterParse[2] and 'STOP' in lineDataAfterParse[3]): 
                    Dstarted = True			
                    isScriptHasD = True
                    endOfexec = True	
        
	    			
	    		# Check the timing sequence
                try:
                    timer = int((re.search('([\\d]+)',lineDataAfterParse[1])).groups()[0])
                except:
                    errorsLog.append('Incorrect Timing sequense ')            		
                    lineNumbers.append(str(counter+11))
                else:
                    if cycleDelay is True:
                        if not (timer-PrevCycleTime) == int(TimeCycle):
                            errorsLog.append('Incorrect Timing sequense ')            		
                            lineNumbers.append(str(counter+11))
        
                    PrevCycleTime = timer
                    cycleDelay = True
        
                # Check the proxy SIL global number (SilGlobReq)
                if len(lineDataAfterParse) >= 6:
                    SilGlobReq_flg = True
                    SilGlobCounte_flg = True				
	    			
                    try:
                        SilGlobReq_temp = int(((re.search('[\s]*([0-9]+)',lineDataAfterParse[3])).groups())[0])
                    except: 
                            pass
                            SilGlobReq_flg = False

                    if SilGlobReq_flg is False:
                        errorsLog.append(' Incorrect global request number')            
                        lineNumbers.append(str(counter+11))
                    elif SilGlobReq_temp == 0:
                        pass
                    elif not (SilGlobReq_temp == SilGlobReq):
                        errorsLog.append('Incorrect global request number')            
                        lineNumbers.append(str(counter+11))
	    				
                # Check the proxy SilGlobCounter sequence
                    try:
                        SilGlobCounter_temp = int(((re.search('[\s]*([0-9]+)',lineDataAfterParse[4])).groups())[0])
                    except: 
                            pass
                            SilGlobCounte_flg = False
        
                    if SilGlobCounte_flg is False:
                        errorsLog.append(' Incorrect global request number')            
                        lineNumbers.append(str(counter+11))
                    elif SilGlobCounter_temp == 0:
                        pass
                    elif (SilGlobCounter_temp - prevCounterValue) is not 1:
                        errorsLog.append('Incorrect SIL global counter sequence')            
                        lineNumbers.append(str(counter+11))
                    prevCounterValue = SilGlobCounter_temp
        
	    				
        
                # check the Spaces in every line after Comma(,) and must start D without any space after time cycle
                counter_2 = 0
                for listIndx in lineDataAfterParse:
                    if '\n' not in listIndx:			
                        counter_2 = counter_2+1
                        if counter_2 == 3:
                            try:
                                exist=(re.match('[.a-zA-Z0-9-_]+$',listIndx)).groups()
                            except:
                                errorsLog.append('Script does not comply with the format <No space><value/timer><comma>')            		
                                lineNumbers.append((counter+11))
                        else:					
                            try:
                                exist=(re.match('[\s\t]*[.a-zA-Z0-9-_]+$',listIndx)).groups()
                            except:
                                errorsLog.append('Script does not comply with the format <one space><value/timer><comma>')            		
                                lineNumbers.append((counter+11))
	    	  
            if Dstarted is False and isScriptHasD is True:
                errorsLog.append('D is missing ')            		
                lineNumbers.append((''+str(counter+11)+', remove this line If it is empty'))
        if endOfexec is False:
            errorsLog.append('Scipt must be ended with C, STOP')
            lineNumbers.append('End line of script file')		
        if isNoScript is True:
            errorsLog.append('There is no test run script in file')
            lineNumbers.append('Null')		
        elif isScriptHasD is False:
            errorsLog.append('Test script does not start with D')	
            lineNumbers.append('Null')        	
    except:
        messagebox.showerror('Error','There is some issue while parsing script, contact masthanvali.s@silver-atena.com ')
        GUIObj.destroy()
        sys.exit()
        
	
def ListTheErrors():
    '''
       Identified Issues in script file will be listed here.
    '''	
	
    fptr= open(fileName[:len(fileName)-4]+'_Errors.txt','w')
    if not len(errorsLog):
        fptr.writelines('\nScript file does not have any issues!!')		
        fptr.close()		
		
        messagebox.showinfo('Information!!','Script file has no issues!!')
        GUIObj.destroy()	

    else:
        for errIndx,errInfo in enumerate(errorsLog):
            strng = '\n'+str(errIndx+1)+' '+str(errInfo)+' (Line number:'+str(lineNumbers[errIndx])+')'+'\n'
            fptr.writelines(strng)
        fptr.close()

    pwd=os.getcwd()+'\\'+fileName[:len(fileName)-4]+'_Errors.txt'
    resInfo = '\nThere are some issues in script, please check issues list in "'+pwd+'"'

    messagebox.showinfo('Information!!',resInfo)
    GUIObj.destroy()	


	
def script_exe(filePath, sciptFileName, TkObject_ref, exeTime, SILreqno):

    global TimeCycle, fileName, SilGlobReq, GUIObj,selectedLoc

    GUIObj = TkObject_ref
    try:
        TimeCycle = int(exeTime)
    except:
        messagebox.showerror('Error','Execution time must be of type Integer')
        GUIObj.destroy()
        sys.exit()
    fileName = sciptFileName
    try:
        SilGlobReq = int(SILreqno)	
    except:
        messagebox.showerror('Error','SIL request number shouldd be of type Integer')
        GUIObj.destroy()
        sys.exit()
	
    selectedLoc = filePath

    Aftersplit = selectedLoc.split('/')
    ActualPath = '/'.join(Aftersplit[:len(Aftersplit)-1])
    os.chdir(ActualPath)
	
    OpenFile() 	
    fileHeaderCheck()
    SignalDescription()
    SyntaxCheckAfterPowerOnDelay()	
    ListTheErrors()