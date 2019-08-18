import os,sys,re
from xlrd import open_workbook
from tkinter import *

def script_exe(flPath,file,TkObject_ref,modulefun,SystemTiming):

    #try:
        tp = []
        #present working dir
        pwd=(re.search('(.*)/.*\..*$',flPath)).groups()[0]
        os.chdir(pwd)
        tp.append(' ')
        tp.append(file)
        tp.append(modulefun)	

        xlsPtr = open_workbook(tp[1])
        HLTP_Template = xlsPtr.sheet_by_index(0)
        numberOfRows = HLTP_Template.nrows
        numberOfCols = HLTP_Template.ncols

        print('\nExecution in Progress...')
		
        if numberOfRows <= 9:
            messagebox.showerror('Error','There is no input and output values in TP')
            TkObject_ref.destroy()			
        if numberOfCols <=2:
            messagebox.showerror('Error','There is no test case columns')
            TkObject_ref.destroy()			

            #SIL based script
            
        if not SystemTiming.isdigit():
            messagebox.showerror('Error','System time should be interger')
            TkObject_ref.destroy()
            sys.exit()
			
        if (int(SystemTiming) == 0):
            messagebox.showerror('Error','System time should not be 0')
            TkObject_ref.destroy()
            sys.exit()
			
        #Count number of "Calibration value" rows
        numberOfCalRows = 0
        for rowIndx in range(10,HLTP_Template.nrows):
            if HLTP_Template.cell(rowIndx,0).value == 'Calibration value':
                numberOfCalRows = numberOfCalRows+1
        fPtr = open(''+(tp[1])[:-4]+'_RTRT_Script_temp.ptu','w')
        
        OneCycledelay = False
        testcaseCounter = 1
        newTest = True
        elementNumber = 1
        colIndex = 2
        dummyInsert = False		
        while numberOfCols > colIndex:
            colFlag = True
            rowFlag = True
              
            if OneCycledelay == True:
                PrevcycleNumber = HLTP_Template.cell(0,colIndex-1).value
                if PrevcycleNumber != HLTP_Template.cell(0,colIndex).value:
                    CurrElement_End = '\n\n\t END ELEMENT  '+str(elementNumber)
                    CurrTestCase_End = '\n\n\tEND TEST  --TEST '+str(testcaseCounter)
                    fPtr.write(CurrElement_End)            
                    fPtr.write(CurrTestCase_End)            
                    testcaseCounter= testcaseCounter+1
                    newTest = True
                else:
                    newTest = False
                    CurrElement_End = '\n\n\t END ELEMENT  '+str(elementNumber)
                    fPtr.write(CurrElement_End)            
                    elementNumber = elementNumber+1
            OneCycledelay = True

            # Dummy Element insertion into script file
            openFileDummy = False
            dumyFile = open('Dummy_dump.txt','w')
            if dummyInsert is True:

                ReadDummyFile = open('Dummy_dump.txt')
                CurrElement_start = '\n\n\t ELEMENT '+str(elementNumber)
                fPtr.write(CurrElement_start)
                fPtr.write('\n\n\t-------Inputs')
                for indx_dummy in ReadDummyFile.readlines():
                    fPtr.write(indx_dummy)
                CurrElement_End = '\n\n\t END ELEMENT  '+str(elementNumber)
                fPtr.write(CurrElement_End)
                dumyFile.close()
                ReadDummyFile.close()
                os.remove('Dummy_dump.txt')
                openFileDummy = True
                dummyInsert = False
                elementNumber = elementNumber+1

            if openFileDummy is True:
                dumyFile = open('Dummy_dump.txt','w')
				
				
            #End of dummy Element

            if newTest == True:
                CurrTestCase_start = '\n\n\tTEST '+str(testcaseCounter)
                fPtr.write(CurrTestCase_start)
                CurrTestCase_Family = '\n\tFAMILY Nominal'
                fPtr.write(CurrTestCase_Family)
                elementNumber = 1
            CurrElement_start = '\n\n\t ELEMENT '+str(elementNumber)
            fPtr.write(CurrElement_start)
            currCol = colIndex    
            for rowIndex in range(7,numberOfRows):
                cellData = HLTP_Template.cell(rowIndex,currCol).value
                # when xls cell type is string/char, then avoid this comparison.
                try:	
                    if cellData.is_integer():
                        cellData = int(cellData)
                except:
                        pass				 
                if rowIndex == 7:
                    if len(cellData):
                        # write comment into script file
                        #print(cellData+'\t'+str(colIndex))
                        fPtr.write('\n\n\t\t'+'Comment {'+'\n')
                        fPtr.write(str(cellData)+''+'\n')
                        fPtr.write('\t\t}')
                if rowIndex == 8:
                    if len(tp) <= 2:
                        print('Must enter three arguments <*.py file> <HLTP.xls> <Maintask()>')
                        exit()
                    numberOfcalls = '\n\n\t\t#'+tp[2]+';'

                    if colIndex+1 == numberOfCols:
                       cellData = SystemTiming
                    else:
                        cellData = abs(HLTP_Template.cell(rowIndex,colIndex).value - HLTP_Template.cell(rowIndex,colIndex+1).value)

                    if int(cellData) > int(SystemTiming):
                        dummyInsert = True
                        numberOfcalls_1='\n\n\t\t--- Function Call\n\t\t'+'#for(i=1;i<='+str(int(int(cellData)/int(SystemTiming))-1)+';i++)'+'\n'+'\t\t#{'+'\n'+'\t\t\t#'+tp[2]+';\n\t\t#}'

						
                if HLTP_Template.cell(rowIndex,0).value == 'Inputs':
                    if len(str(cellData))!=0:
                        if rowFlag == True:
                            fPtr.write('\n\n\t\t----Inputs')
                            rowFlag = False
                        dataItem = '\n\t\tVAR\t%-40s' % str(HLTP_Template.cell(rowIndex,1).value+',')+'\t'+'INIT=%-30s' % (str(cellData)+',')+'\t\t\tEV=INIT'
                    
                        fPtr.write(dataItem)
                        if dummyInsert is True:
                            dumyFile.writelines(dataItem)
						
						
                if HLTP_Template.cell(rowIndex,0).value == 'Outputs':
                    if len(str(cellData))!=0:
                        if colFlag == True:
                            fPtr.write('\n\n\t\t----Outputs')
                            colFlag = False            

                        if elementNumber>1:
                            dataItem = '\n\t\tVAR\t%-40s' % str(HLTP_Template.cell(rowIndex,1).value+',')+'\t'+'INIT=%-20s'%(str(HLTP_Template.cell(rowIndex,currCol-1).value)+',')+'\t\t\t\t\tEV='+str(cellData)
                        else:
                            dataItem = '\n\t\tVAR\t%-40s' % str(HLTP_Template.cell(rowIndex,1).value+',')+'\t'+'INIT==,\t\t\tEV='+str(cellData)
                        fPtr.write(dataItem)

            if dummyInsert is True:
                dumyFile.writelines(numberOfcalls_1) 
            fPtr.write(numberOfcalls)
            colIndex = colIndex+1
        CurrElement_End = '\n\n\t END ELEMENT  '+str(elementNumber)
        CurrTestCase_End = '\n\n\tEND TEST  --TEST '+str(testcaseCounter)
        fPtr.write(CurrElement_End)            
        fPtr.write(CurrTestCase_End)            
            
        fPtr.close()
        #print(''+(tp[1])[:-4]+'_RTRT_Script.ptu)
        fPtr = open(''+(tp[1])[:-4]+'_RTRT_Script_temp.ptu')
        fPtr_ptu = open(''+(tp[1])[:-4]+'_RTRT_Script.ptu','w')
        
        linedata = ''
        commentExist = False
        flag = False
        for lineindex in fPtr.readlines():
            linedata = lineindex
            if lineindex.count('\t\tComment {'):
                linedata = lineindex
                linedata = linedata.replace('{','-----------------------------')
                flag = True
            if commentExist == True:
                if lineindex.count('\t\t}'):
                    linedata = linedata.replace('}','Comment -----------------------------')
                    commentExist = False
                    flag = False
                else:
                    datalist = []
                    datalist.append(linedata)
                    datalist.insert(0,'Comment')
                    linedata = '\t\t'+datalist[0]+' '+datalist[1]
            if flag == True:
                commentExist = True
            fPtr_ptu.write(linedata)
        fPtr.close()    
        fPtr_ptu.close()
        dumyFile.close()
        os.remove((tp[1])[:-4]+'_RTRT_Script_temp.ptu')
        os.remove('Dummy_dump.txt')
		
        msg = '1) Results are generated in '+os.getcwd()+'\\'+(tp[1])[:-4]+'_RTRT_Script.ptu file'\
               '\n\n2) Please edit the generated ptu file and preset the output values'		
        messagebox.showinfo('Information!!',msg)
        TkObject_ref.destroy()	
		
    #except:
    #    messagebox.showerror('Error','While generating results error occored, Please report this problem to masthanvali.s@silver-atena.com')
    #    TkObject_ref.destroy()
		
		
#--------------------------------------------------------------
#1) If u have any inputs with blank data then that input data in script will not exist
#2) This script is for only HMS
#   		t		 t+1			 		t+10 	t+11
# Rep:      1		  1   (dummy 2 to 9)     1       1
#--------------------------------------------------------------