import csv,os,sys,glob
from xlsxwriter import *
from tkinter import *

# This is used to fetch the 'exp' column numbers in csv file

def _getItemPosition(xs, item):
    if isinstance(xs, list):
        for i, it in enumerate(xs):
            for pos in _getItemPosition(it, item):
                yield (i,) + pos
    elif xs == item:
        yield ()

def ResultChecker(csvFilePath,TkObject_ref):
    # read the list of csv files in C:\FormalSILResultChecker\FormalRunFiles
    os.chdir(csvFilePath)
    listOfSCVFiles = glob.glob('*.csv')
    
    # Process only when files exist
    if len(listOfSCVFiles) > 0:
    
        #Report the status in 'Final Result report' xls sheet
        xls = Workbook('FormalRunResult.xls')
        xlsPtr = xls.add_worksheet('Results')
        xlsResCount = 0
    
        for csvFile in listOfSCVFiles:
    
            curWrkingFile = ''
            # Fetch only script name
            temp = csvFile.split('_')[0:len(csvFile.split('_'))-5]
            for ind in temp:
                curWrkingFile = curWrkingFile+ind+'_'

            print('\nResult Analysis in progress for file: '+curWrkingFile)
				
            #Write the status of each file into xls sheet 
            xlsPtr.write(xlsResCount,0,curWrkingFile[0:len(curWrkingFile)-1])
    		
    		#Used to compare each file
            BufText = open('buff.txt','w')
            bufcount = 1
            header = []
            finalStatusCount = 0
            failuresList = ''
              
            #write the csv file data into text file
            with open(csvFile) as f:
                reader = csv.reader(f)
                for row in reader:
                    if bufcount == 14:
                        #find the 'exp' columns positions in result file
                        noOfExpCol = list(_getItemPosition(list(row),'exp'))
                        header = row
    
                    if bufcount>14:
                        seqCounter = row[2]
                        for indx in noOfExpCol:
                            indxNum = indx[0] 
    
                            Exp_val = row[indxNum].lstrip()
                            Act_val = row[indxNum+1].lstrip()
    
                            Exp_val = Exp_val.rstrip()
                            Act_val = Act_val.rstrip()
    
                            if not(Exp_val is 'XXX' or Act_val is 'XXX'):
                                if not(Exp_val == Act_val):
                                    finalStatusCount = finalStatusCount + 1
                                    failuresList = failuresList+'\n Results mis-match at Sequencial counter: '+str(seqCounter)+' and Output variable'+str(header[indxNum+1])
    								
                    bufcount = bufcount+1
    
    
            if finalStatusCount>0:
                xlsPtr.write(xlsResCount,1,'FAIL')
                FinalRes = open(curWrkingFile[0:len(curWrkingFile)-1]+'_ERROR.txt','w')
                FinalRes.writelines(failuresList)
                FinalRes.close()
    
            else:
                xlsPtr.write(xlsResCount,1,'PASS')
    
            BufText.close()
            xlsResCount = xlsResCount+1
    		
            #if xlsResCount == 1:
            #    print('\n\n Results comparision and report generation in progress...')
            #
        os.remove('buff.txt')
        try:
            xls.close()
        except:

            messagebox.showerror('Error','ERROR: Please close FormalRunResult.xls file!!')
            TkObject_ref.destroy()
            sys.exit()
    
        messagebox.showinfo('Results','Results are generated in '+csvFilePath+' with group results sheet FormalRunResult.xls and failure report of each script')
        TkObject_ref.destroy()
        sys.exit()		
    else:
        messagebox.showerror('Error','No csv files in '+csvFilePath+'!!')
        TkObject_ref.destroy()		