import sys,os,re
from xlsxwriter import *
from xlrd import open_workbook
from tkinter import *
from tkinter import messagebox

#if __name__ == '__main__':
def startFun(filepath,fileName,TkObject_ref):
    pwd=(re.search('(.*)/.*\..*$',filepath)).groups()[0]
    os.chdir(pwd)

    xlsName = fileName
    
    #Open the selected workbook
    xlsPtr = open_workbook(xlsName)
    HLTP_Template = xlsPtr.sheet_by_index(0)
    
    #Find number of rows and columns in it
    numberOfRows = HLTP_Template.nrows
    numberOfCols = HLTP_Template.ncols
    # Check selected sheet is empty
    if numberOfRows is 0 or numberOfCols is 0:
        messagebox.showerror('Error','TC Template should not be empty!!')
        sys.exit(0)
        TkObject_ref.destroy()

    if os.path.isfile(xlsName.split('.xl')[0]+'_To_Doors.xlsx') is True:
        try:
            os.remove(xlsName.split('.xl')[0]+'_To_Doors.xlsx')
        except:
            messagebox.showerror('Error','\nPlease close the opened file "'+xlsName.split('.xl')[0]+'_To_Doors.xlsx'+'"\n')
            sys.exit(0)
            TkObject_ref.destroy()			

    #Creating new template from Doors TP
    xls = Workbook(xlsName.split('.xl')[0]+'_To_Doors.xlsx')
    xlsWritePtr = xls.add_worksheet('TC_To_Doors')        


    prevSteNumber = -1
	
    # Write basic information into Excel sheet
    xlsWritePtr.write(0,0,'Object Identifier')
    xlsWritePtr.write(0,1,'Object Number')
    xlsWritePtr.write(0,2,'Object Heading')
    xlsWritePtr.write(0,3,'Object Text')
    xlsWritePtr.write(0,4,'Configuration')
    xlsWritePtr.write(0,5,'Test Category')
    xlsWritePtr.write(0,6,'Test Group')
    xlsWritePtr.write(0,7,'Object Type')
    xlsWritePtr.write(0,8,'SSDD Trace')
    xlsWritePtr.write(0,9,'SSDD Ext Trace')
    xlsWritePtr.write(0,10,'ICD LLR Trace')
    xlsWritePtr.write(0,11,'LLR Trace')
    xlsWritePtr.write(0,12,'Comment')	
    xlsWritePtr.write(0,13,'Temp HLR')
    xlsWritePtr.write(0,14,'Temp LLR')
    xlsWritePtr.write(0,15,'Temp ICD LLR')
    xlsWritePtr.write(0,16,'Notes')
    xlsWritePtr.write(0,17,'Reusable Library Trace')	
	                    
    SetVerify = ''	
    comment = '' 
    testCaseTag = ''
    HLRTraceability = ''
    LLRTraceability = ''
    ICDTraceability = ''
    TestCategory = ''
    configurationItem = ''
	
    writeRowIndx = 1
    for colIndx in range(2,numberOfCols):
	
        if len(str(HLTP_Template.cell(0,colIndx).value)) == 0:
            messagebox.showerror('Error','\n Step number should not be empty at column: '+str(colIndx+1))
            xls.close()
            sys.exit(0)
            TkObject_ref.destroy()
			
        print('\nNumber of columns processed: '+str(colIndx))  
			
        #Check whether row is empty
        emptyColExist = True  
        for rowIndx in range(11,numberOfRows):
            if len(str(HLTP_Template.cell(rowIndx,colIndx).value)) > 0:
                emptyColExist = False
                break

        ## Fetch other information like Time, Comment, LLR traceability, ICD, Test case tag, etc
        if len(str(HLTP_Template.cell(10,colIndx).value)) > 0:
            if len(str(HLTP_Template.cell(9,colIndx).value))>0:
                comment = comment+'At: '+str(HLTP_Template.cell(10,colIndx).value)+' \n'+str(HLTP_Template.cell(9,colIndx).value).strip()+'\n\n'
        else:
            if len(str(HLTP_Template.cell(9,colIndx).value))>0:
                comment = comment+' '+str(HLTP_Template.cell(9,colIndx).value).strip()+'\n\n'
				
        if emptyColExist is False:
            if str(prevSteNumber) == str(HLTP_Template.cell(0,colIndx).value):
                if len(str(HLTP_Template.cell(10,colIndx).value)) > 0:
                    SetVerify = SetVerify + '\nAt: '+ str(HLTP_Template.cell(10,colIndx).value)+'\n'
            
                for rowIndx in range(11,numberOfRows):
                    if (str((HLTP_Template.cell(rowIndx,0).value).strip()) == 'Inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Intermediate inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate Inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip() == 'Intermediate Inputs')):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:
                            SetVerify = SetVerify + 'Set '+str(HLTP_Template.cell(rowIndx,1).value)+' to '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'
                    elif ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Intermediate Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate outputs'):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:				
                            SetVerify = SetVerify + 'Verify '+str(HLTP_Template.cell(rowIndx,1).value)+' is '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'
                    #if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:							
                    #    SetVerify = SetVerify + '\n'
						
                if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:					
                    SetVerify = SetVerify + '\n'
            
                # At the end of column write data into Doors sheet
                if (colIndx+1 == numberOfCols) and len(SetVerify)>1: 
                    xlsWritePtr.write(writeRowIndx,3,SetVerify)
                    xlsWritePtr.write(writeRowIndx,12,comment)
                    comment = ''					
                    SetVerify = ''
                elif (colIndx != numberOfCols):
                    if len(SetVerify)>1 and (HLTP_Template.cell(0,colIndx).value != HLTP_Template.cell(0,colIndx+1).value):
                        xlsWritePtr.write(writeRowIndx,3,SetVerify)
                        xlsWritePtr.write(writeRowIndx,12,comment)
                        comment = ''						
                        SetVerify = ''
                        writeRowIndx = writeRowIndx+1				
                prevSteNumber = str(HLTP_Template.cell(0,colIndx).value)				
            else:

                ## Read test case tag if it exists
                if len(str(HLTP_Template.cell(1,colIndx).value)) > 0:
                    testCaseTag = str(HLTP_Template.cell(1,colIndx).value)

                ## HLR traceability					
                if len(str(HLTP_Template.cell(4,colIndx).value)) > 0:
                    HLRTraceability = str(HLTP_Template.cell(4,colIndx).value)

                ## LLR traceability					
                if len(str(HLTP_Template.cell(5,colIndx).value)) > 0:
                    LLRTraceability = str(HLTP_Template.cell(5,colIndx).value)

                ## ICD traceability					
                if len(str(HLTP_Template.cell(6,colIndx).value)) > 0:
                    ICDTraceability = str(HLTP_Template.cell(6,colIndx).value)
					
                ## Test category
                if len(str(HLTP_Template.cell(7,colIndx).value)) > 0:
                    TestCategory = str(HLTP_Template.cell(7,colIndx).value)

                configurationItem = ''
                ## Read configuration data 
                #for rowInd in range(11,numberOfRows):
                #    if (((str(HLTP_Template.cell(rowInd,0).value).strip()) is 'Calibration value') or ((str(HLTP_Template.cell(rowInd,0).value).strip()) is 'calibration value') or str(((HLTP_Template.cell(rowInd,0).value).strip()) is 'calibration Value') or (((HLTP_Template.cell(rowInd,0).value).strip()) is 'Calibration Value')):
                #        if len(str(HLTP_Template.cell(rowInd,colIndx).value)) > 0:
                #            configurationItem = configurationItem+'Configure '+str(HLTP_Template.cell(rowInd,1).value)+' to '+str(HLTP_Template.cell(rowInd,colIndx).value)+'\n'

                for rowInd in range(11,numberOfRows):
                    try:
                        re.search(r'calibration value',HLTP_Template.cell(rowInd,0).value,re.I).groups()
                        if len(str(HLTP_Template.cell(rowInd,colIndx).value)) > 0:
                            configurationItem = configurationItem+'Configure '+str(HLTP_Template.cell(rowInd,1).value)+' to '+str(HLTP_Template.cell(rowInd,colIndx).value)+'\n'
                    except:pass

                
                if len(testCaseTag)>0:
                    xlsWritePtr.write(writeRowIndx,0,testCaseTag)
                if len(HLRTraceability)>0:
                    xlsWritePtr.write(writeRowIndx,8,HLRTraceability)
                if len(LLRTraceability)>0:
                    xlsWritePtr.write(writeRowIndx,11,LLRTraceability)			
                if len(ICDTraceability)>0:
                    xlsWritePtr.write(writeRowIndx,10,ICDTraceability)			
                if len(TestCategory)>0:
                    xlsWritePtr.write(writeRowIndx,5,TestCategory)			
                if len(configurationItem)>0:		
                    xlsWritePtr.write(writeRowIndx,4,configurationItem)
                
                testCaseTag = ''
                HLRTraceability = ''
                LLRTraceability = ''
                ICDTraceability = ''
                TestCategory = ''
                configurationItem = ''	


                if len(str(HLTP_Template.cell(10,colIndx).value)) > 0:
                    SetVerify = SetVerify + 'At: '+ str(HLTP_Template.cell(10,colIndx).value)+'\n'
            
                for rowIndx in range(11,numberOfRows):
                    if (str((HLTP_Template.cell(rowIndx,0).value).strip()) == 'Inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Intermediate inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate Inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate inputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip() == 'Intermediate Inputs')):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:
                            SetVerify = SetVerify + 'Set '+str(HLTP_Template.cell(rowIndx,1).value)+' to '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'
                    elif ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'Intermediate Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate Outputs') or ((str(HLTP_Template.cell(rowIndx,0).value).strip()) == 'intermediate outputs'):
                        if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:
                            SetVerify = SetVerify + 'Verify '+str(HLTP_Template.cell(rowIndx,1).value)+' is '+str(HLTP_Template.cell(rowIndx,colIndx).value)+'\n'
                if len(str(HLTP_Template.cell(rowIndx,colIndx).value))>0:					
                    SetVerify = SetVerify + '\n'
            
                if (colIndx+1 == numberOfCols) and len(SetVerify)>1: 
                    xlsWritePtr.write(writeRowIndx,3,SetVerify)
                    xlsWritePtr.write(writeRowIndx,12,comment)
                    comment = ''
                    SetVerify = ''           
                # Write data into Doors format sheet
                elif colIndx != numberOfCols:
                    if len(SetVerify)>1 and (HLTP_Template.cell(0,colIndx).value != HLTP_Template.cell(0,colIndx+1).value):
                        xlsWritePtr.write(writeRowIndx,3,SetVerify)
                        xlsWritePtr.write(writeRowIndx,12,comment)
                        comment = ''
                        SetVerify = ''
                        writeRowIndx = writeRowIndx+1
		    			
                prevSteNumber = str(HLTP_Template.cell(0,colIndx).value)			
        else:

            ## Read test case tag if it exists
            if len(str(HLTP_Template.cell(1,colIndx).value)) > 0:
                testCaseTag = str(HLTP_Template.cell(1,colIndx).value)

            ## HLR traceability					
            if len(str(HLTP_Template.cell(4,colIndx).value)) > 0:
                HLRTraceability = str(HLTP_Template.cell(4,colIndx).value)

            ## LLR traceability					
            if len(str(HLTP_Template.cell(5,colIndx).value)) > 0:
                LLRTraceability = str(HLTP_Template.cell(5,colIndx).value)

            ## ICD traceability					
            if len(str(HLTP_Template.cell(6,colIndx).value)) > 0:
                ICDTraceability = str(HLTP_Template.cell(6,colIndx).value)
		
            ## Test category
            if len(str(HLTP_Template.cell(7,colIndx).value)) > 0:
                TestCategory = str(HLTP_Template.cell(7,colIndx).value)

            configurationItem = ''
            ## Read configuration data 
            for rowInd in range(11,numberOfRows):
                try:
                    re.search(r'calibration value',HLTP_Template.cell(rowInd,0).value,re.I).groups()
                    if len(str(HLTP_Template.cell(rowInd,colIndx).value)) > 0:
                        configurationItem = configurationItem+'Configure '+str(HLTP_Template.cell(rowInd,1).value)+' to '+str(HLTP_Template.cell(rowInd,colIndx).value)+'\n'
                except:pass

            if len(comment)>0: 
                xlsWritePtr.write(writeRowIndx,12,HLTP_Template.cell(9,colIndx).value)			
            if len(testCaseTag)>0:
                xlsWritePtr.write(writeRowIndx,0,testCaseTag)			
            if len(HLRTraceability)>0:
                xlsWritePtr.write(writeRowIndx,8,HLRTraceability)
            if len(LLRTraceability)>0:
                xlsWritePtr.write(writeRowIndx,11,LLRTraceability)			
            if len(ICDTraceability)>0:
                xlsWritePtr.write(writeRowIndx,10,ICDTraceability)			
            if len(TestCategory)>0:
                xlsWritePtr.write(writeRowIndx,5,TestCategory)			
            if len(configurationItem)>0:		
                xlsWritePtr.write(writeRowIndx,4,configurationItem)

            SetVerify = ''	
            comment = '' 
            testCaseTag = ''
            HLRTraceability = ''
            LLRTraceability = ''
            ICDTraceability = ''
            TestCategory = ''
            configurationItem = ''				

            writeRowIndx = writeRowIndx+1			
			

    xls.close()
    messagebox.showinfo('Results','Results are generated in '+str(xlsName.split('.xl')[0]+'_To_Doors.xlsx'))
    TkObject_ref.destroy()		
		
'''

1. Ensure that Calibration values are mentioned only the starting of cycle.
2. Template should always be in first sheet

'''