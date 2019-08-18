import sys,os,re
from xlsxwriter import *
from xlrd import open_workbook
from tkinter import *

systemInputs = []
systemOutputs = []
listOfWarnings = []

def ParseSetSection(rowIndx,HLTP_Template):
	
    # Fetch the timings of each cycle

    infoSet = (HLTP_Template.cell(rowIndx,6).value)
    
    SetFormatOfType1 = False
    SetFormatOfType2 = False
    SetFormatOfType3 = False
    
    try:
        SetFormat1 = re.search('At.*cycle.*[\d]+.*for.*[\d]+.*cycles',infoSet,re.M|re.I)
        SetFormat1.group()
        SetFormatOfType1 = True
    except:
        try:
            SetFormat2 = re.search('At.*[\d]+.*cycle.*For.*[\d]+.*cycles',infoSet,re.M|re.I)
            SetFormat2.group()	
            SetFormatOfType2 = True
        except:
            try:
                SetFormat3 = re.search('At:.*[\d]+',infoSet,re.M|re.I)		
                SetFormat3.group()	
                SetFormatOfType3 = True
            except:
                errorIn = 'None of the below set formats match in cell Row Number: '+str(rowIndx+1)+' Column number: 7, Please check and correct it. \n\n1. At cycle <Number> for <Number> cycles \n2. At <Number> cycle For <Number> cycles \n3. At: <Number>'
                
                clearBuffer(errorIn)
                sys.exit(0)

    #Write each cell data into buffer
    
    fWrtPtr = open('BufferSet.txt','w')
    fWrtPtr.write(infoSet)
    fWrtPtr.close()				
    
    #Read data from buffer
    freadPtr = open('BufferSet.txt')

    CompleteInputSet = []    # this will hold the data like [{'ExeutionTime':[1,1],'Input1':'value','Input2':'value',....}] of each cell
    inputSet = {}

    for lineInfo in freadPtr:

        lineInfoAfterStrip = lineInfo.strip()    
		
        #Process the data only when row info is not empty
        if len(lineInfoAfterStrip) > 0:
    	
            # Identify Set section timing
    
            if SetFormatOfType1 is True:
    
                try:
                    SetTime = re.search('At.*cycle.*[\d]+.*for.*[\d]+.*cycles',lineInfoAfterStrip,re.I).group()
                    timing = list(map(int,re.findall(r'\d+',SetTime)))
                    CompleteInputSet.append(inputSet)					
                    inputSet = {}
                    inputSet['ExecutionTime'] = timing
                except:
                    try:
                       inputData = re.search(r'Set(.*) to (.*)',lineInfoAfterStrip,re.I).groups()
                       inputSet[inputData[0].strip()] = inputData[1]
                    except:
                        errorIn = 'While processing cell <row number:'+str(rowIndx+1)+' and column: 7>, one of the below formats does not match \n1.Set(.*) to (.*)\n2. At <Timing formats>'
                        clearBuffer(errorIn)						
                        sys.exit(0)
    
            elif SetFormatOfType2 is True:

                try:
                    SetTime = re.search('At.*[\d]+.*cycle.*For.*[\d]+.*cycles',lineInfoAfterStrip,re.I).group()
                    timing = list(map(int,re.findall(r'\d+',SetTime)))
                    CompleteInputSet.append(inputSet)					
                    inputSet = {}
                    inputSet['ExecutionTime'] = timing
                except:
                    try:
                       inputData = re.search(r'Set(.*) to (.*)',lineInfoAfterStrip,re.I).groups()
                       inputSet[inputData[0].strip()] = inputData[1]
                    except:
                        errorIn = 'While processing cell <row number:'+str(rowIndx+1)+' and column: 7>, one of the below formats does not match \n1.Set(.*) to (.*)\n2. At <Timing formats>'
                        clearBuffer(errorIn)
                        sys.exit(0)	
						
            elif SetFormatOfType3 is True:

                try:
                    SetTime = re.search('At:.*[\d]+',lineInfoAfterStrip,re.I).group()
                    timing = list(map(int,re.findall(r'\d+',SetTime)))
                    CompleteInputSet.append(inputSet)					
                    inputSet = {}
                    inputSet['ExecutionTime'] = timing
                except:
                    try:
                       inputData = re.search(r'Set(.*) to (.*)',lineInfoAfterStrip,re.I).groups()
                       inputSet[inputData[0].strip()] = inputData[1]
                    except:
                        errorIn = 'While processing cell <row number:'+str(rowIndx+1)+' and column: 7>, one of the below formats does not match \n1.Set(.*) to (.*)\n2. At <Timing formats>'
                        clearBuffer(errorIn)
                        sys.exit(0)	

    # Append the finally iterated information to list
    CompleteInputSet.append(inputSet)
    return CompleteInputSet 

def ParseVerifySection(rowIndx,HLTP_Template):

    # Fetch the timings of each cycle

    infoVerify = (HLTP_Template.cell(rowIndx,7).value)
    
    VerifyFormatOfType1 = False
    VerifyFormatOfType2 = False
    VerifyFormatOfType3 = False
    
    try:
        VerifyFormat1 = re.search('At.*the.*end.*of.*cycle.*[\d]+',infoVerify,re.M|re.I)
        VerifyFormat1.group()
        VerifyFormatOfType1 = True
    except:
        try:
            VerifyFormat2 = re.search('At.*the.*end.*of.*[\d]+.*cycle',infoVerify,re.M|re.I)
            VerifyFormat2.group()	
            VerifyFormatOfType2 = True
        except:
            try:
                VerifyFormat3 = re.search('At:.*[\d]+',infoVerify,re.M|re.I)		
                VerifyFormat3.group()	
                VerifyFormatOfType3 = True
            except:
                errorIn = 'None of the below Verify formats match in cell Row Number: '+str(rowIndx+1)+' Column number: 8, Please check and correct it. \n\n1. At the end of <Number> cycle \n2. At the end of cycle <Number> \n3. At: <Number>'
                clearBuffer(errorIn)
                sys.exit(0)

    #Write each cell data into buffer
    
    fWrtPtr = open('BufferVerify.txt','w')
    fWrtPtr.write(infoVerify)
    fWrtPtr.close()				
    
    #Read data from buffer
    freadPtr = open('BufferVerify.txt')

    CompleteInputVerify = []    # this will hold the data like [{'ExeutionTime':[1,1],'Input1':'value','Input2':'value',....}] of each cell
    inputVerify = {}

    for lineInfo in freadPtr:

        lineInfoAfterStrip = lineInfo.strip()    
		
        #Process the data only when row info is not empty
        if len(lineInfoAfterStrip) > 0:
    	
            # Identify Set section timing
    
            if VerifyFormatOfType1 is True:
    
                try:
                    SetTime = re.search('At.*the.*end.*of.*cycle.*[\d]+',lineInfoAfterStrip,re.I).group()
                    timing = list(map(int,re.findall(r'\d+',SetTime)))
                    CompleteInputVerify.append(inputVerify)					
                    inputVerify = {}
                except:
                    try:
                       inputData = re.search(r'Verify(.*) is (.*)',lineInfoAfterStrip,re.I).groups()
                       inputVerify[inputData[0].strip()] = inputData[1]
                    except:
                        errorIn = 'While processing cell <row number:'+str(rowIndx+1)+' and column: 8>, one of the below formats does not match \n1.Verify(.*) is (.*)\n2. At end of <time> '
                        clearBuffer(errorIn)
                        sys.exit(0)
    
            elif VerifyFormatOfType2 is True:

                try:
                    SetTime = re.search('At.*the.*end.*of.*[\d]+.*cycle',lineInfoAfterStrip,re.I).group()
                    timing = list(map(int,re.findall(r'\d+',SetTime)))
                    CompleteInputVerify.append(inputVerify)					
                    inputVerify = {}
                except:
                    try:
                       inputData = re.search(r'Verify(.*) is (.*)',lineInfoAfterStrip,re.I).groups()
                       inputVerify[inputData[0].strip()] = inputData[1]
                    except:
                        errorIn = 'While processing cell <row number:'+str(rowIndx+1)+' and column: 8>, one of the below formats does not match \n1.Verify(.*) is (.*)\n2. At end of <time> '
                        clearBuffer(errorIn)
                        sys.exit(0)	
						
            elif VerifyFormatOfType3 is True:

                try:
                    SetTime = re.search('At:.*[\d]+',lineInfoAfterStrip,re.I).group()
                    timing = list(map(int,re.findall(r'\d+',SetTime)))
                    CompleteInputVerify.append(inputVerify)					
                    inputVerify = {}
                except:
                    try:
                       inputData = re.search(r'Verify(.*) is (.*)',lineInfoAfterStrip,re.I).groups()
                       inputVerify[inputData[0].strip()] = inputData[1]
                    except:
                        errorIn = 'While processing cell <row number:'+str(rowIndx+1)+' and column: 8>, one of the below formats does not match \n1.Verify(.*) is (.*)\n2. At end of <time> '
                        clearBuffer(errorIn)
                        sys.exit(0)	
				
    # Append the finally iterated information to list
    CompleteInputVerify.append(inputVerify)
    return CompleteInputVerify

def ParseCommentSection(rowIndx,HLTP_Template):
    

    infoComment = (HLTP_Template.cell(rowIndx,9).value)
    CompleteInputVerify = []
    CommentInList = []

    #Prase only if cell has information
    if len(infoComment)>0:
	
        CommentFormatOfType1 = False
        CommentFormatOfType2 = False
        CommentFormatOfType3 = False
        TimmingCommentExist = True
        
        try:
            VerifyFormat1 = re.search(r'^\s*At.*Cycle.*[\d]+\s*$',infoComment,re.M|re.I)
            VerifyFormat1.group()
            CommentFormatOfType1 = True
        except:
            try:
                VerifyFormat2 = re.search(r'^\s*At.*[\d]+.*Cycle\s*$',infoComment,re.M|re.I)
                VerifyFormat2.group()	
                CommentFormatOfType2 = True
            except:
                try:
                    VerifyFormat3 = re.search('^\s*At.*[\d]+\s*$',infoComment,re.M|re.I)		
                    VerifyFormat3.group()	
                    CommentFormatOfType3 = True
                except:
                    TimmingCommentExist = False

                    singleComment = {}  
                    singleComment['NoTimingComment'] =  infoComment
                    CompleteInputVerify.append(singleComment)

        if TimmingCommentExist is True:

            fWrtPtr = open('BufferComment.txt','w')        
            fWrtPtr.write(infoComment)
            fWrtPtr.close()

            # Read the data line by line and push it into list
            FreadPtr = open('BufferComment.txt')
            for lineInfo in FreadPtr:
                try:
                    re.search(r'^\s*\n\s*$',lineInfo).group()
                except:  
                    CommentInList.append(lineInfo)
            FreadPtr.close()

            inputString = ''
            timing = 0
            inputVerify = {}

            if CommentFormatOfType1 is True:

                for CommentInListIndx in range(0,len(CommentInList)):
                    lineInfoAfterStrip = CommentInList[CommentInListIndx]

                    #Process the data only when row info is not empty
                    if len(lineInfoAfterStrip) > 0:

                        # Identify Set section timing
                        try:                
                            SetTime = re.search('^\s*At.*Cycle.*[\d]+\s*$',lineInfoAfterStrip,re.I).group()

                            timing = list(map(int,re.findall(r'\d+',SetTime)))

                        except:
                            if CommentInListIndx < len(CommentInList):
                                try:
                                    re.search('^\s*At.*Cycle.*[\d]+\s*$',CommentInList[CommentInListIndx+1],re.I).group()
                                    inputString = inputString + lineInfoAfterStrip
                                    if timing != 0:
                                        inputVerify['ExecutionTime'] = timing[0]
                                        inputVerify['Comment'] = inputString

                                        CompleteInputVerify.append(inputVerify)

                                    inputString = ''
                                    inputVerify = {}
                                except:  
                                    inputString = inputString + lineInfoAfterStrip

                if timing != 0:
                    inputVerify['ExecutionTime'] = timing[0]
                    inputVerify['Comment'] = inputString
                    
                    CompleteInputVerify.append(inputVerify)

            if CommentFormatOfType2 is True:

                for CommentInListIndx in range(0,len(CommentInList)):
                    lineInfoAfterStrip = CommentInList[CommentInListIndx]

                    #Process the data only when row info is not empty
                    if len(lineInfoAfterStrip) > 0:

                        # Identify Set section timing
                        try:                
                            SetTime = re.search('^\s*At.*[\d]+.*Cycle\s*$',lineInfoAfterStrip,re.I).group()

                            timing = list(map(int,re.findall(r'\d+',SetTime)))

                        except:
                            if CommentInListIndx < len(CommentInList):
                                try:
                                    re.search('^\s*At.*[\d]+.*Cycle\s*$',CommentInList[CommentInListIndx+1],re.I).group()
                                    inputString = inputString + lineInfoAfterStrip
                                    if timing != 0:
                                        inputVerify['ExecutionTime'] = timing[0]
                                        inputVerify['Comment'] = inputString

                                        CompleteInputVerify.append(inputVerify)

                                    inputString = ''
                                    inputVerify = {}
                                except:  
                                    inputString = inputString + lineInfoAfterStrip

                if timing != 0:
                    inputVerify['ExecutionTime'] = timing[0]
                    inputVerify['Comment'] = inputString
                    
                    CompleteInputVerify.append(inputVerify)	

            if CommentFormatOfType3 is True:

                for CommentInListIndx in range(0,len(CommentInList)):
                    lineInfoAfterStrip = CommentInList[CommentInListIndx]

                    #Process the data only when row info is not empty
                    if len(lineInfoAfterStrip) > 0:

                        # Identify Set section timing
                        try:                
                            SetTime = re.search('^\s*At.*[\d]+\s*$',lineInfoAfterStrip,re.I).group()

                            timing = list(map(int,re.findall(r'\d+',SetTime)))

                        except:
                            if CommentInListIndx < len(CommentInList):
                                try:
                                    re.search('^\s*At.*[\d]+\s*$',CommentInList[CommentInListIndx+1],re.I).group()
                                    inputString = inputString + lineInfoAfterStrip
                                    if timing != 0:
                                        inputVerify['ExecutionTime'] = timing[0]
                                        inputVerify['Comment'] = inputString

                                        CompleteInputVerify.append(inputVerify)

                                    inputString = ''
                                    inputVerify = {}
                                except:  
                                    inputString = inputString + lineInfoAfterStrip

                if timing != 0:
                    inputVerify['ExecutionTime'] = timing[0]
                    inputVerify['Comment'] = inputString
                    
                    CompleteInputVerify.append(inputVerify)						

        return  CompleteInputVerify
	
def writeHeaderIntoTemplate(xlsWritePtr):
  
    #Write header into xls sheet
	
    xlsWritePtr.write(0,0,'SSDD v x.xx')
    xlsWritePtr.write(0,1,'TP#')
	
    xlsWritePtr.write(1,0,'Test Procedure Tag/ Obj ID')

    xlsWritePtr.merge_range('A3:B3','Set')
    xlsWritePtr.merge_range('A4:B4','Verify')
    xlsWritePtr.merge_range('A5:B5','Configure')	
    xlsWritePtr.merge_range('A6:B6','Test case Trace')
    xlsWritePtr.merge_range('A7:B7','Configuration')
    xlsWritePtr.merge_range('A8:B8','Comment')
	
	
	
    xlsWritePtr.merge_range('A9:B9','Time ( ms)')	
	
    xlsWritePtr.set_column('A:B', 20)
	

def writeDataIntoTemplate(xlsWritePtr,ProcedureStepInfo):

    global globalInputxlsxWriter
    global globalOutputxlsxWriter
    global newInputRow 	
    global newOutputRow	
    global systemInputs 	
    global systemOutputs

    TimingComment = False

    # Write the data into template

    xlsWritePtr.write(1,globalInputxlsxWriter, ProcedureStepInfo['Doors ID'])

    # delete the dummy element from Set
    del ProcedureStepInfo['Set'][0]

    xlsWritePtr.write(5,globalInputxlsxWriter, ProcedureStepInfo['TC Trace'])
    xlsWritePtr.write(6,globalInputxlsxWriter, ProcedureStepInfo['Configuration'])

    commentOfCell = ProcedureStepInfo['Comment']

    # Write commment into template
    if commentOfCell is not None:
        if 'NoTimingComment' in commentOfCell[0].keys():
            xlsWritePtr.write(7,globalInputxlsxWriter, commentOfCell[0]['NoTimingComment'])
        else:
            TimingComment = True
	
    # Fetch each cell Set data
    for eachSetinfo in ProcedureStepInfo['Set']:
	
        if len(str(ProcedureStepInfo['Step Number']))>0:
            xlsWritePtr.write(0,globalInputxlsxWriter, int(ProcedureStepInfo['Step Number']))		
        else:
            xlsWritePtr.write(0,globalInputxlsxWriter, ProcedureStepInfo['Step Number'])

        # Write execution time into template		
        timeingofCell = eachSetinfo['ExecutionTime'][0]
        xlsWritePtr.write(8,globalInputxlsxWriter, timeingofCell)		
		
        # Write commment into template
        if TimingComment is True:
            for commentIndx in commentOfCell:
                if int(timeingofCell) == int(commentIndx['ExecutionTime']):
                    xlsWritePtr.write(7,globalInputxlsxWriter,commentIndx['Comment'])

        #Fetch each input item
        for eachInputInfo in list(eachSetinfo.items()):

            if eachInputInfo[0] != 'ExecutionTime':

                #check if input item is exist in list
                if eachInputInfo[0] in systemInputs:
                    xlsWritePtr.write(systemInputs.index(eachInputInfo[0])+9,0,'Inputs')
                    xlsWritePtr.write(systemInputs.index(eachInputInfo[0])+9,globalInputxlsxWriter, eachInputInfo[1])
                else:
                    xlsWritePtr.write(newInputRow,0,'Inputs')				
                    systemInputs.append(eachInputInfo[0])
                    xlsWritePtr.write(newInputRow,1, eachInputInfo[0])						
                    xlsWritePtr.write(newInputRow,globalInputxlsxWriter, eachInputInfo[1])
                    newInputRow = newInputRow+1

        globalInputxlsxWriter = globalInputxlsxWriter+1
		
		
    # delete the dummy element from Verify
    del ProcedureStepInfo['Verify'][0]
	
    # Fetch each cell Verify data
    for eachSetinfo in ProcedureStepInfo['Verify']:
	
        #Fetch each input item
        for eachInputInfo in list(eachSetinfo.items()):


            #check if input item is exist in list
            
            if eachInputInfo[0] in systemOutputs:
                xlsWritePtr.write(systemOutputs.index(eachInputInfo[0])+500,0,'Outputs')
                xlsWritePtr.write(systemOutputs.index(eachInputInfo[0])+500,globalOutputxlsxWriter, eachInputInfo[1])
            else:
                xlsWritePtr.write(newOutputRow+500,0,'Outputs')				
                systemOutputs.append(eachInputInfo[0])
                xlsWritePtr.write(newOutputRow+500,1, eachInputInfo[0])						
                xlsWritePtr.write(newOutputRow+500,globalOutputxlsxWriter, eachInputInfo[1])
                newOutputRow = newOutputRow+1

        globalOutputxlsxWriter = globalOutputxlsxWriter+1    

def deleteUnwantedRows():
        # Delete unwanted rows from xlsheet

        #Open the selected workbook
        xlsPtr2 = open_workbook(xlsName.split('.xl')[0]+'_Template_Buff.xlsx')

        #Write data into Final template 
        xls2 = Workbook(xlsName.split('.xl')[0]+'_Template.xlsx')				

        nSheets = xlsPtr2.nsheets
        sheetNames = xlsPtr2.sheets()

        for nsheetIndx in range(0,nSheets):		
		
			#read data from sheet
            HLTP_Template2 = xlsPtr2.sheet_by_index(nsheetIndx)

			#Find number of rows and columns in it
            numberOfRows = HLTP_Template2.nrows
            numberOfCols = HLTP_Template2.ncols

            print('\n Recheck of '+sheetNames[nsheetIndx].name+' in progress...')
			
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
def ParseSetVerifyAndCommentInputData(rowIndx,xlsWritePtr,HLTP_Template):

    ProcedureStepInfo = {}
    ProcedureStepInfo['Doors ID'] = HLTP_Template.cell(rowIndx,0).value
    ProcedureStepInfo['Step Number'] = HLTP_Template.cell(rowIndx,5).value
    ProcedureStepInfo['Configuration'] = HLTP_Template.cell(rowIndx,8).value
    ProcedureStepInfo['TC Trace'] = HLTP_Template.cell(rowIndx,10).value

    # Parse the Set section data	
    SetSectionData = ParseSetSection(rowIndx,HLTP_Template)
    ProcedureStepInfo['Set'] = SetSectionData	
	
    #Parse the Verify section data
    VerifySectionData = ParseVerifySection(rowIndx,HLTP_Template)
    ProcedureStepInfo['Verify'] = VerifySectionData
    
    if len(SetSectionData) != len(VerifySectionData): 
        errorIn = '\nNumber of inputs Set and Outputs Verify sections does not match in row:'+str(rowIndx+1)+' column:<7,8>'
        clearBuffer(errorIn)
        sys.exit(0)	

    #Parse the Comment section
    CommentSectionData = ParseCommentSection(rowIndx,HLTP_Template)
    ProcedureStepInfo['Comment'] = CommentSectionData

    #Write Data into template
    writeDataIntoTemplate(xlsWritePtr,ProcedureStepInfo)
	
def clearBuffer(errorInfo):
    global listOfWarnings
    listOfWarnings.append(errorInfo)
	
def startFun(filepath,fileName,TkObject_ref):
    pwd=(re.search('(.*)/.*\..*$',filepath)).groups()[0]
    os.chdir(pwd)

    global xlsName
    global globalInputxlsxWriter
    global globalOutputxlsxWriter
    global newInputRow 	
    global newOutputRow	
    global systemInputs 	
    global systemOutputs

    xlsName = fileName
    
    
    #Open the selected workbook
    xlsPtr = open_workbook(xlsName)
    HLTP_Template = xlsPtr.sheet_by_index(0)
    
    #Find number of rows and columns in it
    numberOfRows = HLTP_Template.nrows
    numberOfCols = HLTP_Template.ncols
    
    # Check selected sheet is empty
    if numberOfRows is 0 or numberOfCols is 0:
        messagebox.showerror('Error','Doors TP should not be empty!!')
        TkObject_ref.destroy()
        sys.exit(0)
    
    elif numberOfCols < 10:
        messagebox.showerror('Error','Selected Doors TP should have following fields \n\n1.ID\n2.Object Number\n3.Object Heading\n4.Procedure Title\n5.Object Type\n6.Step\n7.Set\n8.Verify\n9.Configuration\n10.Comments\n11.TC Trace ')
        TkObject_ref.destroy()
        sys.exit(0)
    	
    #Find total HL procedures ['.*HL_TEST.*']
    procedures = []
    
    for indx in range(0,numberOfRows):
        if ('HL_TEST' in HLTP_Template.cell(indx,3).value) or ('HL_Test' in HLTP_Template.cell(indx,3).value):
            procedures.append(indx)
			
    #If 'HL test' are available in Doors TP, then only process furthur
    if len(procedures) is 0:
        messagebox.showerror('Error','There is no HL Test procedures in selected TP')
        TkObject_ref.destroy()		
        sys.exit(0)
    
    try:
        #Creating new template from Doors TP
        xls = Workbook(xlsName.split('.xl')[0]+'_Template_Buff.xlsx')
        
    except:
        messagebox.showerror('Error','\nPlease close the opened file "'+xlsName.split('.xl')[0]+'_Template_Buff.xlsx'+'"\n')
        TkObject_ref.destroy()		
        sys.exit(0)	

    for HLTestrowNumber in range(0,len(procedures)):
 
        print('\n\n Processing of xlsx sheet "'+HLTP_Template.cell(procedures[HLTestrowNumber],3).value+'" in progress....')
    
        #add a sheet to xls with name of HLTP procedure
        xlsWritePtr = xls.add_worksheet(HLTP_Template.cell(procedures[HLTestrowNumber],3).value)
    
        globalInputxlsxWriter = 2
        globalOutputxlsxWriter = 2
    
        newInputRow = 9
        newOutputRow = 0
        systemInputs = []
        systemOutputs = []		
    
        #Write header information into Excel sheet
        writeHeaderIntoTemplate(xlsWritePtr)
		
        #Loop through all rows till next HLTP
        if (len(procedures)-1) == HLTestrowNumber:
            maxRowIter = numberOfRows
        else:
            maxRowIter = procedures[HLTestrowNumber+1]
    	
        string1 = 'Procedure Step'
        rowwrtIndx = 0
        procStepexist = False
    	
        for rowIndx in range(procedures[HLTestrowNumber],maxRowIter):
    
            try:
                re.search(r'Procedure.*Step',HLTP_Template.cell(rowIndx,4).value,re.I).group()
				
                #Invoke user defined Parser function
                ParseSetVerifyAndCommentInputData(rowIndx,xlsWritePtr,HLTP_Template)
    
                rowwrtIndx = rowwrtIndx+1
                procStepexist = True
				
            except:
                if ((len(procedures)-1) == HLTestrowNumber) and (procStepexist is True):
                    #At the end of HL TEST procedure stop processing
                    break
					
                try:
                    re.search(r'Procedure.*Step',HLTP_Template.cell(rowIndx,4).value,re.I).group()
                except:
                    if (procedures[HLTestrowNumber] == procedures[len(procedures)-1]) and (procStepexist is True):
                        break

    xls.close()

    # Delete unwanted rows from buffer	
    deleteUnwantedRows()
    
    #Delete files
    try:
        os.remove('BufferSet.txt')
        os.remove('BufferVerify.txt')
        os.remove('BufferComment.txt')
        
        os.remove(xlsName.split('.xl')[0]+'_Template_Buff.xlsx')
    except:
        pass
    
    if len(listOfWarnings)>0:
    
        ErrorPtr = open('DoorsTPIssues.txt','w')
        for errInfo in listOfWarnings:
            ErrorPtr.write(errInfo+'\n\n')
        ErrorPtr.close()
    
        messagebox.showinfo('Issues in Doors TP','\n\nThere are some issues in selected Doors TP, errors are reported in DoorsTPIssues.txt file')
    
    if len(listOfWarnings)>0:
        messagebox.showinfo('Issues in Doors TP','\n\n1. Results are generated in '+xlsName[0:len(xlsName)-5]+'_Template.xlsx. \n\n2. Since there are some warnings, template might not have correct results. Please correct listed errors')
        TkObject_ref.destroy()		
    else:
        messagebox.showinfo('Results','\n Results are generated in '+xlsName[0:len(xlsName)-5]+'_Template.xlsx\n')	
        TkObject_ref.destroy()