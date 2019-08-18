import sys,os,re
from xlsxwriter import *
from xlrd import open_workbook
from tkinter import *
from tkinter import messagebox

#if __name__ == '__main__':
def startFun(filepath,fileName,TkObject_ref,exeTime):
    pwd=(re.search('(.*)/.*\..*$',filepath)).groups()[0]
    os.chdir(pwd)

    xlsName = fileName
	
    try:
        systemExecTime = int(exeTime)
    except:
        messagebox.showerror('Error','System execution time should be of type Integer!!')
        sys.exit()
        TkObject_ref.destroy()
    
    #Open the selected workbook
    xlsPtr = open_workbook(xlsName)
    HLTP_Template = xlsPtr.sheet_by_index(0)
    
    #Find number of rows and columns in it
    numberOfRows = HLTP_Template.nrows
    numberOfCols = HLTP_Template.ncols
    # Check selected sheet is empty
    if numberOfRows is 0 or numberOfCols is 0:
        messagebox.showerror('TC Template should not be empty!!')
        sys.exit(0)
        TkObject_ref.destroy()

    if os.path.isfile(xlsName.split('.xl')[0]+'_To_Doors.xlsx') is True:
        try:
            os.remove(xlsName.split('.xl')[0]+'_To_Doors.xlsx')
        except:
            messagebox.showerror('Error - Close opened file','\nPlease close the opened file "'+xlsName.split('.xl')[0]+'_To_Doors.xlsx'+'"\n')
            sys.exit(0)
            TkObject_ref.destroy()			

    #Creating new template from Doors TP
    xls = Workbook(xlsName.split('.xl')[0]+'_To_Doors.xlsx')
    xlsWritePtr = xls.add_worksheet('TP_To_Doors')        


    prevSteNumber = -1
	
    # Write basic information into Excel sheet
    xlsWritePtr.write(0,0,'Object Identifier')
    xlsWritePtr.write(0,1,'Object Number')
    xlsWritePtr.write(0,2,'Object Heading')
    xlsWritePtr.write(0,3,'Procedure Title')
    xlsWritePtr.write(0,4,'Object Type')
    xlsWritePtr.write(0,5,'Step')
    xlsWritePtr.write(0,6,'Set')
    xlsWritePtr.write(0,7,'Verify')
    xlsWritePtr.write(0,8,'Configuration')
    xlsWritePtr.write(0,9,'Comments')
    xlsWritePtr.write(0,10,'TC Trace')
    xlsWritePtr.write(0,11,'Test Method')	
	                    
    setInfo = ''	
    verifyInfo = ''		
    comment = '' 
    testProcedureTag = ''
    TestCaseraceability = ''
    TestCategory = ''
    configurationItem = ''
	
    writeRowIndx = 1
    for colIndx in range(2,numberOfCols):
        if len(str(HLTP_Template.cell(0,colIndx).value)) == 0:
            messagebox.showerror('\n Step number should not be empty at column: '+str(colIndx+1))
            xls.close()
            sys.exit(0)
            TkObject_ref.destroy()
			
        print('\nNumber of columns processed: '+str(colIndx))  
			
        #Check whether row is empty
        emptyColExist = True  
        for rowIndx in range(9,numberOfRows):
            if len(str(HLTP_Template.cell(rowIndx,colIndx).value)) > 0:
                emptyColExist = False
                break
				
        if emptyColExist is False:
            if str(prevSteNumber) == str(HLTP_Template.cell(0,colIndx).value):
			
                # Calculate time
                if len(str(HLTP_Template.cell(8,colIndx).value)) > 0:
                    try:
                        timeCycle = int(HLTP_Template.cell(8,colIndx).value)/systemExecTime
                    except:
                        messagebox.showerror('Time Error','Time cycle must be of type Integer at Column: '+str(colIndx+1))
                        xls.close()
                        sys.exit()
                        TkObject_ref.destroy()						
                    
                    if(colIndx == numberOfCols-1):
                        setInfo = setInfo + 'At cycle '+str(int(timeCycle))+' For 1 cycles\n'
                        verifyInfo = verifyInfo + 'At the end of cycle '+ str(int(timeCycle))+'\n'
                        comment = comment + 'At the end of cycle '+ str(int(timeCycle))+'\n'						
                    else:
                        try:
                            timeCyclenext = int(HLTP_Template.cell(8,colIndx+1).value)/systemExecTime
                        except:
                            messagebox.showerror('Time Error','Time cycle must be of type Integer at Column: '+str(colIndx+2))
                            xls.close()							
                            sys.exit()
                            TkObject_ref.destroy()
							
							
                        setInfo = setInfo + 'At cycle '+str(int(timeCycle))+' For '+str(int(timeCyclenext-timeCycle))+' cycles\n'
                        verifyInfo = verifyInfo + 'At the end of cycle '+ str(int(timeCycle))+'\n'
                        comment = comment + 'At the end of cycle '+ str(int(timeCycle))+'\n'												

            
                for rowIndx in range(9,numberOfRows):
                    if (str((HLTP_Template.cell(rowIndx,0).value).strip()) == 'Inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'inputs'):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:
                            if(len(str(HLTP_Template.cell(rowIndx,colIndx).value).split('.0'))) > 1:
                                setInfo = setInfo + 'Set '+str(HLTP_Template.cell(rowIndx,1).value)+' to '+str(int(HLTP_Template.cell(rowIndx,colIndx).value))+'\n'
                            else:
                                setInfo = setInfo + 'Set '+str(HLTP_Template.cell(rowIndx,1).value)+' to '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'							
                    elif ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'outputs'):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:
                            if(len(str(HLTP_Template.cell(rowIndx,colIndx).value).split('.0'))) > 1:
                                verifyInfo = verifyInfo + 'Verify '+str(HLTP_Template.cell(rowIndx,1).value)+' is '+str(int(HLTP_Template.cell(rowIndx,colIndx).value))+'\n'
                            else:
                                verifyInfo = verifyInfo + 'Verify '+str(HLTP_Template.cell(rowIndx,1).value)+' is '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'

                if len(setInfo.strip()) > 0: 							
                    setInfo = setInfo + '\n'

                if len(verifyInfo.strip()) > 0: 
                    verifyInfo = verifyInfo + '\n'

					
                ## Comment
                if len(str(HLTP_Template.cell(7,colIndx).value)) > 0:
                    comment = comment+ str(HLTP_Template.cell(7,colIndx).value).strip()+'\n\n'					
					
                ## TC traceability					
                if len(str(HLTP_Template.cell(5,colIndx).value)) > 0:
                    TestCaseraceability = TestCaseraceability+ str(HLTP_Template.cell(5,colIndx).value).strip()+'\n'
					
                if (colIndx+1 > numberOfCols-1): 
                    xlsWritePtr.write(writeRowIndx,6,setInfo)					
                    xlsWritePtr.write(writeRowIndx,7,verifyInfo)
                    xlsWritePtr.write(writeRowIndx,10,TestCaseraceability.strip())	
                    xlsWritePtr.write(writeRowIndx,9,comment.strip())						
                    TestCaseraceability = ''					
                    verifyInfo = ''	
                    comment = ''
                    writeRowIndx = writeRowIndx+1					

                elif (str(HLTP_Template.cell(0,colIndx).value) != str(str(HLTP_Template.cell(0,colIndx+1).value))) and (colIndx != numberOfCols-1): 
                    xlsWritePtr.write(writeRowIndx,7,verifyInfo)
                    xlsWritePtr.write(writeRowIndx,6,setInfo)	
                    xlsWritePtr.write(writeRowIndx,10,TestCaseraceability.strip())		
                    xlsWritePtr.write(writeRowIndx,9,comment.strip())						
                    TestCaseraceability = ''					
                    setInfo = ''									
                    verifyInfo = ''					
                    comment = ''					
                    writeRowIndx = writeRowIndx+1					

                prevSteNumber = str(HLTP_Template.cell(0,colIndx).value)				
            else:

                ## Read test case tag if it exists
                if len(str(HLTP_Template.cell(1,colIndx).value)) > 0:
                    testProcedureTag = str(HLTP_Template.cell(1,colIndx).value)

                ## HLR traceability					
                if len(str(HLTP_Template.cell(5,colIndx).value)) > 0:
                    TestCaseraceability = TestCaseraceability+ str(HLTP_Template.cell(5,colIndx).value).strip()+'\n'

                ## Test category
                if len(str(HLTP_Template.cell(6,colIndx).value)) > 0:
                    TestCategory = str(HLTP_Template.cell(6,colIndx).value)

                if len(testProcedureTag)>0:
                    xlsWritePtr.write(writeRowIndx,0,testProcedureTag)
                if len(TestCategory)>0:
                    xlsWritePtr.write(writeRowIndx,8,TestCategory)			
                
                testProcedureTag = ''
                TestCategory = ''

                # Calculate time
                if len(str(HLTP_Template.cell(8,colIndx).value)) > 0:
                    try:
                        timeCycle = int(HLTP_Template.cell(8,colIndx).value)/systemExecTime
                    except:
                        messagebox.showerror('Time Error','Time cycle must be of type Integer at Column: '+str(colIndx+1))
                        xls.close()						
                        sys.exit()
                        TkObject_ref.destroy()						
                    
                    if(colIndx == numberOfCols-1):
                        setInfo = setInfo + 'At cycle '+str(int(timeCycle))+' For 1 cycles\n'
                        verifyInfo = verifyInfo + 'At the end of cycle '+ str(int(timeCycle))+'\n'
                        comment = comment + 'At the end of cycle '+ str(int(timeCycle))+'\n'						
                    else:
                        try:
                            timeCyclenext = int(HLTP_Template.cell(8,colIndx+1).value)/systemExecTime
                        except:
                            messagebox.showerror('Time Error','Time cycle must be of type Integer at Column: '+str(colIndx+2))
                            xls.close()							
                            sys.exit()
                            TkObject_ref.destroy()

                        setInfo = setInfo + 'At cycle '+str(int(timeCycle))+' For '+str(int(timeCyclenext-timeCycle))+' cycles\n'
                        verifyInfo = verifyInfo + 'At the end of cycle '+ str(int(timeCycle))+'\n'
                        comment = comment + 'At the end of cycle '+ str(int(timeCycle))+'\n'												
            
                for rowIndx in range(9,numberOfRows):
                    if (str((HLTP_Template.cell(rowIndx,0).value).strip()) == 'Inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'inputs'):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:
                            if(len(str(HLTP_Template.cell(rowIndx,colIndx).value).split('.0'))) > 1:
                                setInfo = setInfo + 'Set '+str(HLTP_Template.cell(rowIndx,1).value)+' to '+str(int(HLTP_Template.cell(rowIndx,colIndx).value))+'\n'
                            else:
                                setInfo = setInfo + 'Set '+str(HLTP_Template.cell(rowIndx,1).value)+' to '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'							
                    elif ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'outputs'):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:
                            if(len(str(HLTP_Template.cell(rowIndx,colIndx).value).split('.0'))) > 1:
                                verifyInfo = verifyInfo + 'Verify '+str(HLTP_Template.cell(rowIndx,1).value)+' is '+str(int(HLTP_Template.cell(rowIndx,colIndx).value))+'\n'
                            else:
                                verifyInfo = verifyInfo + 'Verify '+str(HLTP_Template.cell(rowIndx,1).value)+' is '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'

                ## Comment
                if len(str(HLTP_Template.cell(7,colIndx).value)) > 0:
                    comment = comment+ str(HLTP_Template.cell(7,colIndx).value).strip()+'\n\n'	
							
                if len(setInfo.strip()) > 0: 							
                    setInfo = setInfo + '\n'

                if len(verifyInfo.strip()) > 0: 
                    verifyInfo = verifyInfo + '\n'
            
                if (colIndx+1 > numberOfCols-1): 
                    xlsWritePtr.write(writeRowIndx,6,setInfo)				
                    xlsWritePtr.write(writeRowIndx,7,verifyInfo)
                    xlsWritePtr.write(writeRowIndx,10,TestCaseraceability.strip())					
                    xlsWritePtr.write(writeRowIndx,9,comment.strip())
                    verifyInfo = ''	
                    TestCaseraceability = ''					
                    comment = ''
                    writeRowIndx = writeRowIndx+1					

                elif (str(HLTP_Template.cell(0,colIndx).value) != str(str(HLTP_Template.cell(0,colIndx+1).value))) and (colIndx != numberOfCols-1): 
                    xlsWritePtr.write(writeRowIndx,7,verifyInfo)
                    xlsWritePtr.write(writeRowIndx,6,setInfo)	
                    xlsWritePtr.write(writeRowIndx,10,TestCaseraceability.strip())					
                    xlsWritePtr.write(writeRowIndx,9,comment.strip())					
                    setInfo = ''									
                    verifyInfo = ''					
                    TestCaseraceability = ''
                    comment = ''					
                    writeRowIndx = writeRowIndx+1					
					
                prevSteNumber = str(HLTP_Template.cell(0,colIndx).value)			
        else:

            ## Read test case tag if it exists
            if len(str(HLTP_Template.cell(1,colIndx).value)) > 0:
                testProcedureTag = str(HLTP_Template.cell(1,colIndx).value)

            ## HLR traceability					
            if len(str(HLTP_Template.cell(5,colIndx).value)) > 0:
                TestCaseraceability = str(HLTP_Template.cell(5,colIndx).value)

            ## Test category
            if len(str(HLTP_Template.cell(6,colIndx).value)) > 0:
                TestCategory = str(HLTP_Template.cell(6,colIndx).value)
				
            

            if len(testProcedureTag)>0:
                xlsWritePtr.write(writeRowIndx,0,testProcedureTag)
            if len(TestCaseraceability)>0:
                xlsWritePtr.write(writeRowIndx,10,TestCaseraceability)
            if len(TestCategory)>0:
                xlsWritePtr.write(writeRowIndx,8,TestCategory)			

            setInfo = ''
            verifyInfo = ''			
            comment = '' 
            testProcedureTag = ''
            TestCaseraceability = ''
            TestCategory = ''

            writeRowIndx = writeRowIndx+1			
			

    xls.close()
    messagebox.showinfo('Results','Results are generated in '+str(xlsName.split('.xl')[0]+'_To_Doors.xlsx'))
    TkObject_ref.destroy()		
		
'''

1. Ensure that Calibration values are mentioned only the starting of cycle.
2. Template should always be in first sheet

'''