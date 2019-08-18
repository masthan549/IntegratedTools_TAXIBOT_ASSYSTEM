import re,sys,os
from xlrd import open_workbook
from tkinter import *
from tkinter import messagebox


def GenerateTS(filepath,filepath_temp,TkObject_ref,sheetIndx,systemExeTime,SILGlobNumber,SILScriptStartTime):
    '''
	   This function converts the TS into SIL TS
	'''

############# Enter these details ################ 

    try:
        XlsName = filepath_temp
        sysExecTime = int(systemExeTime)
        sheetIndx = int(sheetIndx)-1
        GlobalReqNumber = float(SILGlobNumber)
        seqCounter = 1.0
        ScriptStartTime = int(SILScriptStartTime)
        rowNumber = 9
    except:
        messagebox.showerror('Error','May be you have entered incorrect values in required fields')
        TkObject_ref.destroy()		
#############################


    pwd=(re.search('(.*)/.*\..*$',filepath)).groups()[0]
    os.chdir(pwd)

    try: 
        XlsName_Modified = XlsName.split('.xl')
	    
        #Open xls sheet 
	    
        fptr = open_workbook(XlsName)
        HLTP_Template = fptr.sheet_by_index(sheetIndx)
        rows = HLTP_Template.nrows
        cols = HLTP_Template.ncols
    except:
        messagebox.showerror('Error','Error occured while opening selected file')
        TkObject_ref.destroy()		
		
    try:	
        #open text file
        os.chdir(os.getcwd())
        SILTS = open(XlsName_Modified[0]+'_HL_TEST.txt','w')
        IssuesPtr = open('TP_To_SIL_Script_Issues.txt','w')
	    
        print('\nExecution in Progress...')  
        
        NextCellData = False		
        #Read the data from xls and push it into text file
        for indx in range(2,HLTP_Template.ncols):
            rowInfo = ''
            timeDiffInCycles = 0
            ColInfo = HLTP_Template.col_values(indx)
            ColInfo = ColInfo[8:len(ColInfo)]
            curCycleTime = int(ColInfo[0])
        
            #Write Template data into text file only when there is a input or output
        
            # Calculate timing information
            if indx+1 != HLTP_Template.ncols:
	    	
                if sysExecTime>=10:
                    ColInfo2 = HLTP_Template.col_values(indx+1)
                    ColInfo2 = ColInfo2[8:len(ColInfo2)]
                    timeDiff = int(ColInfo2[0]) - curCycleTime
                else:
                    ColInfo2 = HLTP_Template.col_values(indx+1)
                    ColInfo2 = ColInfo2[8:len(ColInfo2)]
                    timeDiff = int(ColInfo2[0]) - curCycleTime
            else:
                if sysExecTime>=10:
                    timeDiff = sysExecTime
                else:
                    timeDiff = sysExecTime/sysExecTime			
        
	    	
            if NextCellData is False:				
                rowInfo = ' 1, '+str(ScriptStartTime)+',D, '+str(float(GlobalReqNumber))+', '+str(seqCounter)+','
            else:
                rowInfo = ' 1, '+str(ScriptStartTime)+',D, '+str(float(GlobalReqNumber))+', '+str(seqCounter)+','
        
            colInfoCopy = ''
            for indx2 in range(1,len(ColInfo)):
                if(str((HLTP_Template.cell(indx2+8,0).value)).strip() != ''):
                    try:
                        rowInfo = rowInfo+' '+str(float(ColInfo[indx2])).strip()+','
                        colInfoCopy = colInfoCopy+' '+str(float(ColInfo[indx2])).strip()+','
                    except:
                        IssuesPtr.write('In column number:'+str(indx+1)+' given value is: '+str(ColInfo[indx2].strip())+', but it should be of type integer or float.\n')
                        rowInfo = rowInfo+' '+str(ColInfo[indx2].strip())+','
                        colInfoCopy = colInfoCopy+' '+str(ColInfo[indx2].strip())+','
            SILTS.write(rowInfo+'\n')
        
        
            if sysExecTime >=10:
                if timeDiff > sysExecTime:
                    for indx3 in range(0,int(timeDiff/sysExecTime)-1):
                        seqCounter = seqCounter+1				
                        ScriptStartTime = sysExecTime+ScriptStartTime				
                        rowInfo = ''
                        rowInfo = ' 1, '+str(ScriptStartTime)+',D, '+str(float(GlobalReqNumber))+', '+str(seqCounter)+','+colInfoCopy
                        SILTS.write(rowInfo+'\n') 	
                ScriptStartTime = sysExecTime+ScriptStartTime							
            elif timeDiff > (sysExecTime/sysExecTime):
                for indx3 in range(0,timeDiff-1):
                    seqCounter = seqCounter+1			
                    ScriptStartTime = (sysExecTime*10)+ScriptStartTime				
                    rowInfo = ''
                    rowInfo = ' 1, '+str(ScriptStartTime)+',D, '+str(float(GlobalReqNumber))+', '+str(seqCounter)+','+colInfoCopy
                    SILTS.write(rowInfo+'\n') 					
        
                ScriptStartTime = (sysExecTime*10)+ScriptStartTime		
        
            seqCounter = seqCounter+1
        
        SILTS.close() 
        IssuesPtr.close()
	    
        SILPtrOpen = open('TP_To_SIL_Script_Issues.txt')
        dataExist = 0
        #check whether data exist in file
        for InfoInFile in SILPtrOpen.readlines():
            if InfoInFile != '':
                dataExist = dataExist+1
	    		
        if dataExist>0:
            messagebox.showinfo('Information!!','Some issues found during test procedure to script conversion. Issue are written into TP_To_SIL_Script_Issues')
            messagebox.showinfo('Information!!','Results are written into '+XlsName_Modified[0]+'_HL_TEST.txt'+'')			
            SILPtrOpen.close()
            TkObject_ref.destroy()
            sys.exit()
        else:
            messagebox.showinfo('Information!!','Results are written into '+XlsName_Modified[0]+'_HL_TEST.txt'+'')			
            SILPtrOpen.close()
            TkObject_ref.destroy()
            os.remove('TP_To_SIL_Script_Issues.txt')
            sys.exit()
    except:
        messagebox.showerror('Error','1) You are may be not following the process/syntax, so please refer provided procedure for tool run \n\n 2) If you still have any issues, then report masthanvali.s@silver-atena.com ')
        TkObject_ref.destroy()	
		
		
'''
1. From HLTP test Procedure delete Calibration values rows (keep only Input and Output rows).
2. Please ensure that there is no content after outputs
3. If TP is cycle based, then enter execution time between 1 to 9 (depends on the time diff). If TP is timing based, then min execution time is 10.
4. TP template should be as configured
'''		