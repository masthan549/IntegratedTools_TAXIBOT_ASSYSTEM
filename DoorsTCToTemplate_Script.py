import sys,os,re
from xlsxwriter import *
from xlrd import open_workbook
from tkinter import *
import DoorsTCToTemplate_GUI
from tkinter import messagebox

userActionIsRequired = []
systemInput = []
systemOutput = []
systemCal = []

global inputRowCount
global outputRowCount
global calRowCount


def writeHeaderIntoTemplate():

    xlsTemplatePtr.write(0,0,'SSDD v x.xx')
    xlsTemplatePtr.write(0,1,'TC#')
	
    xlsTemplatePtr.merge_range('A2:B2','Test Case Tag')
    xlsTemplatePtr.merge_range('A3:B3','Test Case')
    xlsTemplatePtr.merge_range('A4:B4','Configure')
    xlsTemplatePtr.merge_range('A5:B5','HLR Traceability')	
    xlsTemplatePtr.merge_range('A6:B6','LLR Traceability')
    xlsTemplatePtr.merge_range('A7:B7','ICD LLR Traceability')
    xlsTemplatePtr.merge_range('A8:B8','Test category')	
    xlsTemplatePtr.merge_range('A9:B9','config/values')	
    xlsTemplatePtr.merge_range('A10:B10','Comment')	
    xlsTemplatePtr.merge_range('A11:B11','Time ( cycles)')	

    xlsTemplatePtr.set_column('A:B', 20)	

def checkHeaderSequence():
    
    headerFormat = '1. Object Identifier, \n2. Object Number, \n3. Object Heading, \n4. Object Text, \n5. Configuration, \n6. Test Category, \n7. Test Group, \n8. Object Type, \n9. SSDD Trace, \n10. SSDD Ext Trace, \n11. ICD LLR Trace, \n12. LLR Trace, \n13. Comment, \n14. Temp HLR, \n15. Temp LLR, \n16. Temp ICD LLR, \n17. Notes, \n18. Reusable Library Trace'
	
    if not (((xlsReaderPtr.cell(0,0).value) == 'Object Identifier') and ((xlsReaderPtr.cell(0,3).value) == 'Object Text') and ((xlsReaderPtr.cell(0,4).value) == 'Configuration') and ((xlsReaderPtr.cell(0,5).value  == 'Test Category')) and ((xlsReaderPtr.cell(0,6).value) == 'Test Group') and ((xlsReaderPtr.cell(0,7).value) == 'Object Type') and ((xlsReaderPtr.cell(0,8).value) == 'SSDD Trace') and ((xlsReaderPtr.cell(0,9).value) == 'SSDD Ext Trace') and ((xlsReaderPtr.cell(0,10).value) == 'ICD LLR Trace') and ((xlsReaderPtr.cell(0,11).value) == 'LLR Trace') and ((xlsReaderPtr.cell(0,12).value) == 'Comment')):
        messagebox.showerror('Error','Test case document Header format is not as below \n'+headerFormat)
        TkObject_ref.destroy()
        sys.exit()

		
def timingDataWriteIntoTemplate(timing_row_info,indx_val):

    global colCount

    # Timing value analysis
    timingNumberInLst = []
    numberOfTiming = re.findall(r'(At[\s]*:[\s]*t[[\s]*[\\+]*[\s]*[0-9]*]*)',timing_row_info,re.I|re.M)
    for lstIndx in numberOfTiming:
        tempList = []
        extContent = re.search(r'At[\s]*:[\s]*(.*)',lstIndx.strip(),re.I|re.M).groups()[0]
        tempList.append(lstIndx.strip())		
        tempList.append(str(extContent))
        timingNumberInLst.append(tempList);

    #Find the position of time in completeList
    completeList = timing_row_info.split('\n')
    afterTrimList = []
    for lstVal in completeList:
        afterTrimList.append(lstVal.strip())

    for key_val in range(0,len(timingNumberInLst)):
        timeIndx = afterTrimList.index(timingNumberInLst[key_val][0])
        timingNumberInLst[key_val].append(timeIndx)

    indx_time = 0
    initDelay = TRUE
    for indx_time_2 in range(0,len(afterTrimList)):
        if afterTrimList[indx_time_2] != '':

            if (int(timingNumberInLst[indx_time][2]) == indx_time_2):

                if initDelay is FALSE:
                    colCount = colCount + 1

                xlsTemplatePtr.write(10,colCount,timingNumberInLst[indx_time][1])
				
                indx_time = indx_time+1	
                if len(timingNumberInLst) == indx_time:
                    indx_time = len(timingNumberInLst)-1
				
			
                # Write Set and Verify data into template

                isSetTrue = False
                # Extract the inputs
                try:
                    InSetgroup = re.search(r'Set[\s]+(.*)[\s]+to[\s]+(.*)',afterTrimList[indx_time_2],re.I)
                    InList = InSetgroup.groups()
                
                    # Find whether input variable already exist in input table
                    findInOutVariableInBuff_Template(TRUE,InList[0].strip(),InList[1].strip())
                    isSetTrue = True
                except:
                    pass # Set section is identified
                
                if isSetTrue is False:
                    # Extract the outputs
                    try:
                        OutSetgroup = re.search(r'Verify[\s]+(.*)[\s]+is[\s]+(.*)',afterTrimList[indx_time_2],re.I)
                        OutList = OutSetgroup.groups()
                
                        # Find whether input variable already exist in input table
                        findInOutVariableInBuff_Template(FALSE,OutList[0].strip(),OutList[1].strip())
                    except:
                        try:
                            re.match(r'^At.*',afterTrimList[indx_time_2],re.I).groups()
                        except: 
                            userActionIsRequired.append('At row number: '+str(indx_val+1)+' any one of the following formats does not match \na. Set (.*) to (.*)\nb. Verify (.*) is (.*)\n\n')						
            
            
            else:
			
			    # Write the step number into template
                xlsTemplatePtr.write(0,colCount,indx_val)			
			
                initDelay = FALSE      
            
                # Write Set and Verify data into template
            
                isSetTrue = False
                # Extract the inputs
                try:
                    InSetgroup = re.search(r'Set[\s]+(.*)[\s]+to[\s]+(.*)',afterTrimList[indx_time_2],re.I)
                    InList = InSetgroup.groups()
                
                    # Find whether input variable already exist in input table
                    findInOutVariableInBuff_Template(TRUE,InList[0].strip(),InList[1].strip())
                    isSetTrue = True
                except: pass
                    # Set section is identified
                
                if isSetTrue is False:
                    # Extract the outputs
                    try:
                        OutSetgroup = re.search(r'Verify[\s]+(.*)[\s]+is[\s]+(.*)',afterTrimList[indx_time_2],re.I)
                        OutList = OutSetgroup.groups()
                
                        # Find whether input variable already exist in input table
                        findInOutVariableInBuff_Template(FALSE,OutList[0].strip(),OutList[1].strip())
                    except:
                        try:
                            re.match(r'^At.*',afterTrimList[indx_time_2],re.I).groups()
                        except: 
                            userActionIsRequired.append('At row number: '+str(indx_val+1)+' any one of the following formats does not match \na. Set (.*) to (.*)\nb. Verify (.*) is (.*)\n\n')						
		    	
		
def Parse_ObjectText_Config_Info(indx_val):

    global calRowCount
    global systemCal
    
    # Check for Time cycles
    rowInfo = xlsReaderPtr.row_values(indx_val)

    #Calibration values placed in template
    calibInfo = rowInfo[4].strip()
    if calibInfo != '': 
        for indx_cal in rowInfo[4].split('\n'):
            if indx_cal.strip() != '':
                cal_info = re.match(r'Configure[\s]+(.*)[\s]+to[\s]+(.+)',indx_cal.strip(),re.I)
                exceptionRaised = False 
                try:
                    cal_var_val = cal_info.groups()
                except:
                    userActionIsRequired.append('At row number: '+str(indx_val+1)+' Configuration values should be in the format of <Configure <var name> to <value>> \n\n')						
                    exceptionRaised = True

                if exceptionRaised is False:
                    #Write the calib data into template
                    if cal_var_val[0].strip() in systemCal:
                        xlsTemplatePtr.write(systemCal.index(cal_var_val[0].strip())+11,colCount,cal_var_val[1].strip())
                    else:
                        xlsTemplatePtr.write(calRowCount,0,'Calibration value')
                        xlsTemplatePtr.write(calRowCount,1,cal_var_val[0].strip())			
                        xlsTemplatePtr.write(calRowCount,colCount,cal_var_val[1].strip())
                        systemCal.append(cal_var_val[0].strip())			
                        calRowCount = calRowCount+1 

	
    # strip the white spaces and newlines at the beginning or ending of the string
    objectText = rowInfo[3].strip()
	 
    # Alternative1: Object Text
    try:
	
        Format1 = re.search(r'At[\s]*:[\s]*t.*',objectText,re.I|re.M)
        Format1.group()
    
        try:
            # Parse the test cases and write into template
            timingDataWriteIntoTemplate(objectText,indx_val)
        except: 
            userActionIsRequired.append('At row number: '+str(indx_val+1)+' there is a problem at parsing the time cycles. Make sure timing cycle in the format of "At: <time cycle>\nSet <input> to <value>\nVerify<output> to <value>"\n\n')		

    except:
        FormatIsSet = FALSE

        try:
		
		    #Alternative2: Set and Verify alternative case

            # Check for whether it is Verification case or Normal case 
            Format2 = re.match(r'^Set.*',objectText,re.I)
            Format2.group()

            #write a logic to split the test cases into cells
            
            tempList = objectText.split('\n')
            for InOutSet in tempList:
                isSetTrue = False
                # Extract the inputs
                try:
				
			        # Write the step number into template
                    xlsTemplatePtr.write(0,colCount,indx_val)					
				
                    InSetgroup = re.search(r'Set[\s]+(.*)[\s]+to[\s]+(.*)',InOutSet,re.I)
                    InList = InSetgroup.groups()

                    # Find whether input variable already exist in input table
                    findInOutVariableInBuff_Template(TRUE,InList[0].strip(),InList[1].strip())
                    isSetTrue = True
                except:
                    pass # Set section is identified

                if isSetTrue is False:
                    # Extract the outputs
                    try:
                        OutSetgroup = re.search(r'Verify[\s]+(.*)[\s]+is[\s]+(.*)',InOutSet,re.I)
                        OutList = OutSetgroup.groups()

                        # Find whether input variable already exist in input table
                        findInOutVariableInBuff_Template(FALSE,OutList[0].strip(),OutList[1].strip())
						
                    except:
                        userActionIsRequired.append('At row number: '+str(indx_val+1)+' any one of the following formats does not match \na. Set (.*) to (.*)\nb. Verify (.*) is (.*)\n\n')

        except:
		
		    #Alternative3: All verification cases
			
			# Write the step number into template
            xlsTemplatePtr.write(0,colCount,indx_val)
			
            # write all verification cases into comment section
            xlsTemplatePtr.write(9,colCount,objectText)


def findInOutVariableInBuff_Template(InputVar,varName,valVal):
    #find the input or output variable position in buffer

    global systemInput
    global systemOutput
    global colCount
    global inputRowCount
    global outputRowCount

    if InputVar is TRUE:
        #Find Input
        if varName in systemInput:
            xlsTemplatePtr.write(systemInput.index(varName)+170,colCount,valVal)
        else:
            xlsTemplatePtr.write(inputRowCount,0,'Inputs')
            xlsTemplatePtr.write(inputRowCount,1,varName)			
            xlsTemplatePtr.write(inputRowCount,colCount,valVal)
            systemInput.append(varName)			
            inputRowCount = inputRowCount+1
    else:
        #Find Output
        if varName in systemOutput:
            xlsTemplatePtr.write(systemOutput.index(varName)+463,colCount,valVal)
        else:
            xlsTemplatePtr.write(outputRowCount,0,'Outputs')
            xlsTemplatePtr.write(outputRowCount,1,varName)			
            xlsTemplatePtr.write(outputRowCount,colCount,valVal)
            systemOutput.append(varName)			
            outputRowCount = outputRowCount+1
    			
def Parse_tescaseInfo():
    
    #Check for Test case Header sequence
    checkHeaderSequence()
    global colCount
    colCount = 2

    #Fet the object text and parse it
    for indx in range(1,numberOfRows):
        rowInfo = xlsReaderPtr.row_values(indx)
  
        if (rowInfo[3].strip()).find('DELETE') is not 0:

            # write HLTP number
            if len(rowInfo[0]) > 0:
                xlsTemplatePtr.write(1,colCount,rowInfo[0])
		    	
            # Write SSDD number
            if len(rowInfo[8]) > 0: #IAI
                xlsTemplatePtr.write(4,colCount,rowInfo[8])
            elif len(rowInfo[9]) > 0: #Ricordo
                xlsTemplatePtr.write(4,colCount,rowInfo[9])
            
            # write ICD LLR
            if len(rowInfo[10]) > 0:
                xlsTemplatePtr.write(6,colCount,rowInfo[10])
            
            # write LLR trace
            if len(rowInfo[11]) > 0:
                xlsTemplatePtr.write(5,colCount,rowInfo[11])
            
            # write test category
            if len(rowInfo[5]) > 0:
                xlsTemplatePtr.write(7,colCount,rowInfo[5])
		    	
            # Write test group into template
            if len(rowInfo[6]) > 0:
                xlsTemplatePtr.write(8,colCount,rowInfo[6])
		    
            # Write the comment information into template section
            if len(rowInfo[12]) > 0:
                xlsTemplatePtr.write(9,colCount,rowInfo[12])        
		    
            #Parse the Object Text and Configuration information
            Parse_ObjectText_Config_Info(indx)
            colCount = colCount+1

def deleteUnwantedRows():
        # Delete unwanted rows from xlsheet

        #Open the selected workbook
        xlsPtr2 = open_workbook(xlsxTemplate_Buff)

        #Write data into Final template 
        xls2 = Workbook(xlsxTemplate)				

        nSheets = xlsPtr2.nsheets
        sheetNames = xlsPtr2.sheets()

        for nsheetIndx in range(0,nSheets):		
		
			#read data from sheet
            HLTP_Template2 = xlsPtr2.sheet_by_index(nsheetIndx)

			#Find number of rows and columns in it
            numberOfRows = HLTP_Template2.nrows
            numberOfCols = HLTP_Template2.ncols

	        #write data into sheet
            xlsWritePtr2 = xls2.add_worksheet(sheetNames[nsheetIndx].name+' ')
            writeRowIndx = 0
            writeColIndx = 0
	        
            for rowIndx in range(0,numberOfRows):
                for ColIndx in range(0,numberOfCols):
                    if len(HLTP_Template2.cell(rowIndx,0).value) > 2:
                        xlsWritePtr2.write(writeRowIndx,writeColIndx,HLTP_Template2.cell(rowIndx,ColIndx).value)
                        writeColIndx = writeColIndx+1
                if len(HLTP_Template2.cell(rowIndx,0).value) > 2:
                    writeRowIndx = writeRowIndx+1
                    writeColIndx = 0
            
        xls2.close()		

def startFun(filepath,selectedXlsFile,TkObject_ref):


    pwd=(re.search('(.*)/.*\..*$',filepath)).groups()[0]
    os.chdir(pwd)

    global inputRowCount
    global outputRowCount
    global calRowCount
	
    calRowCount = 11 # range is 12 to 61 (maximum holds 50 calibration values)	
    inputRowCount = 170 # range is 62 to 462 (maximum holds 400 calibration values)	
    outputRowCount = 463 # range is 463 onwards	
	
    #Create a new xlsx or xls Template
    global xlsxTemplate
    global xlsxTemplate_Buff	
	
    xlsxTemplate = ''
    xlsxTemplate_Buff = ''	
	
    if(selectedXlsFile.endswith('.xlsx')):
        xlsxTemplate = selectedXlsFile[0:len(selectedXlsFile)-5]+'_Template.xlsx'
        xlsxTemplate_Buff = selectedXlsFile[0:len(selectedXlsFile)-5]+'_Template_Buff.xlsx'		
    elif(selectedXlsFile.endswith('.xls')):
        xlsxTemplate = selectedXlsFile[0:len(selectedXlsFile)-5]+'_Template.xls'
        xlsxTemplate_Buff = selectedXlsFile[0:len(selectedXlsFile)-5]+'_Template_Buff.xls'		
    else:
        messagebox.showerror('Error','Select only xls/xlsx file!!')
        TkObject_ref.destroy()			
        sys.exit()

	# Delete the selected file if it is exist in selected path
    if os.path.isfile(xlsxTemplate):
        try:
            os.remove(xlsxTemplate)
        except:
            messagebox.showerror('Error','Please close opened Excel file: '+xlsxTemplate+' and run it again')
            TkObject_ref.destroy()	
            sys.exit()
	
    # Open a Excel sheet with the name of '<selectedfile>_Template'
    global xlsTemplatePtr		
    xlsTemplate = Workbook(xlsxTemplate_Buff)
    xlsTemplatePtr = xlsTemplate.add_worksheet('HLTC_Template')


    print('Execution in Progress.................\n')
	
    #Write header information into template
    writeHeaderIntoTemplate()

    # Open selected Workbook
    global xlsReaderPtr
    xlsReader = open_workbook(selectedXlsFile)
    xlsReaderPtr = xlsReader.sheet_by_index(0)
	
    #Find number of rows and columns in it
    global numberOfCols
    global numberOfRows
	
	
    numberOfRows = xlsReaderPtr.nrows
    numberOfCols = xlsReaderPtr.ncols	
	
    # Check selected sheet is empty
    if numberOfRows is 0 or numberOfCols is 0:	
        messagebox.showerror('Error','Doors TC should not be empty!!')
        TkObject_ref.destroy()		
        sys.exit()
	
    #Read the data from Doors test case Xls
    Parse_tescaseInfo()

	#Close all opened files
    xlsTemplate.close()

    # Delete unwanted rows from buffer	
    deleteUnwantedRows()    
	
    # Clear buffer
    os.remove(xlsxTemplate_Buff)

    # print warnings list	
    if len(userActionIsRequired) > 0:
	
        fPtr = open(selectedXlsFile[0:len(selectedXlsFile)-5]+'.txt','w')   
        count = 1
	
        for indx in userActionIsRequired:
            fPtr.writelines(str(count)+'. '+indx)
            count = count+1
        fPtr.close()
        messagebox.showinfo('Results', 'Results are generated in :'+str(xlsxTemplate)+', but there are some issues in selected file, please correct them. Issues are written into text file: '+str(selectedXlsFile[0:len(selectedXlsFile)-5]+'.txt'))		
    else:
        messagebox.showinfo('Results','Results are generated in :'+str(xlsxTemplate)+' without any issues')
		
    TkObject_ref.destroy()
		
'''		
 1. Wherever Set and Verify format does not match, then error row info will be written into usetAction list and Test Case Object number will not be written into template
 2. Comment section for test case is not separtaed with time cycles, so entire comment section will be written into the initial step of the each test cse
 3. Set and Verify and Configure formats
    Set <Var name> to <value>
	Verify <Var Name> is <Value>
	Configure <var name> is <Value>
 4. File should be in format of *.xls or *.xlsx
 5. Template file should (from which it is selected) not open
 6. Doors file should have first sheet with the name of 'HLTC_Template' and it should not be empty
 7. Issues which are found in selected sheet will be listed in '<Selected xlsx file>.txt' file
 8. Before time cycles of the test case there should not be any information
'''

