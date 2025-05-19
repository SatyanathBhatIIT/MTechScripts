#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Apr 26 13:47:10 2022

@author: divya
"""
import sqlite3
from sqlite3 import Error
import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import date
import numpy as np

def getYearCurr():
    yearCurrentLastTwo = date.today().year%100
    return yearCurrentLastTwo

def generateYearCols():
    templateArray = ['GATE%dRollNo', 'GATE%dRank', 'GATE%dScore', 'GATE%dDisc']
    templateType = ['Roll','Rank', 'Score', 'Disc']
    yearAll = [0,-1,-2]
    yearCurr = getYearCurr()
    
    yearCols = []
    for currYear in yearAll:
        currYearAbs = yearCurr + currYear
        for curridx, currTemplate in enumerate(templateArray):
            currString = currTemplate % (currYearAbs, )
            yearCols.append((currString,templateType[curridx]))
    return yearCols

def isInList(coap, listOfIssues):
    foundInList = False
    reason = ''
    if len(listOfIssues) !=0:
        if coap in listOfIssues:
            foundInList = True
            reason = listOfIssues[coap]
        
    return foundInList, reason
        

def getDecision(coap, clarityList, rejectList):
    decision = 'Y'
    reason = ''
    
    foundInList, reason = isInList(coap, rejectList)   
    if foundInList == True:
        decision = 'N'
    else:
        foundInList, reason = isInList(coap, clarityList)
        if foundInList == True:
            decision = '?'
    return decision, reason

#modeOfRun = 'CSE'  # EE or CE or ME possible
#modeOfRun = 'CSE'
#modeOfRun = 'EE'
#modeOfRun = 'EC'
modeOfRun = 'MST'
#create file names appropriately
rawData = ''
filteredData = ''
originalData = ''
coapData = ''
programCode = []
sheetName = ''
extraRow = True

if modeOfRun == 'CSE':
    rawData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_CSE_final_deduplicated.xlsx'
    filteredData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_CSE.xlsx'
    originalData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_CSE_original.xlsx'
    coapData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/COAPData_CSE.xlsx'
    programCode = ['CS']
    sheetName = 'CSE.xls'
    sheetName = 'Sheet1'
    extraRow = False
elif modeOfRun == 'CE':
    rawData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/CE.xlsx'
    filteredData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_CE.xlsx'
    originalData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_CE_original.xlsx'
    coapData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/COAPData_CE.xlsx'
    programCode = ['CH', 'ME', 'XE', 'MT', 'BT']
    sheetName = 'CE.xls'
elif modeOfRun == 'ME':
    rawData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/ME.xlsx'
    filteredData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_ME.xlsx'
    originalData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_ME_original.xlsx'
    coapData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/COAPData_ME.xlsx'
    programCode = ['ME']
    sheetName = 'ME.xls'
elif modeOfRun == 'EE':
    rawData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/With manual edits EE.xlsx'
    filteredData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_EE.xlsx'
    originalData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_EE_original.xlsx'
    coapData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/COAPData_EE.xlsx'
    programCode = ['EE']
    sheetName = 'EE.xls'
elif modeOfRun == 'EC':
    rawData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/With manual edits EE.xlsx'
    filteredData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_EC.xlsx'
    originalData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_EC_original.xlsx'
    coapData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/COAPData_EC.xlsx'
    programCode = ['EC']
    sheetName = 'EE.xls'
elif modeOfRun == 'EEnEC':    
    rawData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/With manual edits EE.xlsx'
    filteredData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_EEnEC.xlsx'
    originalData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_EEnEC_original.xlsx'
    coapData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/COAPData_EEnEC.xlsx'
    programCode = ['EE', 'EC']
    sheetName = 'EE.xls'
elif modeOfRun == 'ME':
    rawData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/ME.xlsx'
    filteredData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_ME.xlsx'
    originalData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/MasterData_ME_original.xlsx'
    coapData ='/home/satyanath/work/MTech2025/MtechFirstRoundScript/MTech2025/COAPData_ME.xlsx'
    programCode = ['ME']
    sheetName = 'ME.xls'
elif modeOfRun == 'MST':
    rawData ='/home/satyanath/data/MTech2025/MST.xlsx'
    filteredData ='/home/satyanath/data/MTech2025/MasterData_MST.xlsx'
    originalData ='/home/satyanath/data/MTech2025/MasterData_MST_original.xlsx'
    coapData ='/home/satyanath/data/MTech2025/COAPData_MST.xlsx'
    programCode = ['EE','ME','EC','IN','NM','CE','CS','AE']
    sheetName = 'MST.xls'


#DONE




wb = load_workbook(filename = rawData)
df = pd.DataFrame(wb[sheetName].values)

if extraRow == True:
    df.columns = df.iloc[2]
    df = df[3:]
else:
    df.columns = df.iloc[0]
    df = df[1:]    


column_name_list = df.columns.tolist()
column_name_list[-1] = 'unnamed'
if extraRow == False:
    column_name_list[-2] = 'unnamed2'
    column_name_list[0] = 'invalid'
df.columns= column_name_list


disqualified_list ={}
qualified_list ={}
ask_for_clarity = {}
validDegrees = ['BTech', 'BE',  'MCA', 'MSc(Engg)', 'BSc(Engg)']
gen_ews_obc = ['GEN', 'OBC']
threshold_gen = 60
sc_st = ['SC', 'ST']
threshold_sc = 55
COLUMN_NAMES_OLD= ['Mtech Application No',	'GATE Reg No (without papercode)','COAP','GATE Score']
COLUMN_NAMES = ['Application Seq No','App Status','Remarks','App Date (dd/MMM/yyyy)','Mtech Application No','GATE Reg No (without papercode)','GATE  papercode','GATE Score','Institute Name','Institute ID', 'Shortlisted','Reason']
dfOutput = pd.DataFrame(columns=COLUMN_NAMES) # Note that there are now row data inserted.
df['GATE Roll num'] =""
yearCols  = generateYearCols()
#create an empty dataframe to store filtered data
filtered_rows = []
for index, row in df.iterrows():
    #curr_MAXGATEscore = row['MAX GATE score']

    curr_Appno = row['App no']
    curr_COAP = row['COAP']
    #TBD: curr_GATERollnum = row[ 'GATE Roll num']
    curr_Email = row['Email']
    curr_FullName = row['Full Name']
    curr_Admcat = row['Adm cat']
    curr_Pwd = row['Pwd']
    curr_Ews = row['Ews']
    curr_Gender= row['Gender']
    curr_Category = row['Category']
    
    currGateScoreAll  = []
    currGateScoreRollNumAll = []
        
    currGateScoreDisc = []

    curr_yearCategories = []
    for yearCol in yearCols:
        currVal = row[yearCol[0]]
        curr_yearCategories.append(currVal)
        if yearCol[1]=='Roll':
            currGateScoreRollNumAll.append(currVal or "")
        elif yearCol[1] == "Score":
            currGateScoreAll.append(currVal or "")
        elif yearCol[1] == 'Disc':
            currGateScoreDisc.append(currVal or "")
        
        
    
 
    curr_MaxGATEScore3 = row['MaxGATEScore out of 3 yrs']
    #print(curr_MaxGATEScore3)
    #TBD : curr_MAXGATEscoreFormula = row['MAX GATE score calculated'] 
    #TBD : curr_MAXGATEscoreRollNum_formula = row['MAX GATE score Roll Num']
    curr_HSSCper = row['HSSC(per)']       

    curr_HSSCboard = row['HSSC(board)']
    curr_HSSCdate = row['HSSC(date)']
    curr_SSCboard = row['SSC(board)']
    curr_SSCdate = row['SSC(date)']
    curr_SSCper = row['SSC(per)']
    curr_DegreeQual = row['Degree(Qualification)']
    curr_DegreePassingDate = row['Degree(PassingDate)']
    curr_DegreeBranch = row['Degree(Branch)']
    curr_DegreeOtherBranch = row['Degree(OtherBranch)']
    curr_DegreeInstituteName = row['Degree(Institute Name)']
    curr_DegreeCGPA_7thSem = row['Degree(CGPA-7thSem)']
    curr_DegreeCGPA_8thSem = row['Degree(CGPA-8thSem)']
    curr_DegreePer_7thSem = row['Degree(Per-7thSem)']
    curr_DegreePer_8thSem = row['Degree(Per-8thSem)']
    curr_unnamed = row['unnamed']
    
    
    if curr_DegreeQual not in validDegrees:
        #print(curr_COAP + ' '+ curr_Email + ' '+ curr_DegreeQual)
        if (curr_DegreeQual == 'MSc'):
            ask_for_clarity[curr_COAP]= ( curr_Email, 'Unclear if candidate holds MSc(Engg) degree')            
    currPerThresh = threshold_gen    
    if curr_Category in gen_ews_obc:
        pass
    elif(curr_Category in sc_st):
        currPerThresh = threshold_sc
    else:
        ask_for_clarity[curr_COAP]= (curr_Email, 'Category not entered ',)        
    if curr_Pwd=='Yes':
        currPerThresh = threshold_sc
        #print (curr_COAP + ' PWD ' + curr_Email)

    
    curr_CPIThresh = currPerThresh*0.1    
    currPerc = -100.0 ## invalid percentage
    
    if (curr_DegreePer_8thSem==None and curr_DegreePer_7thSem==None):
    ## look for CPI
        curr_CPI = 0.0
        if (curr_DegreeCGPA_8thSem == None and curr_DegreeCGPA_7thSem == None):
            print(curr_COAP + 'no CGPA or Per')
            #disqualified_list
        elif(curr_DegreeCGPA_8thSem==None):
            curr_CPI= curr_DegreeCGPA_7thSem
        else:
            curr_CPI = curr_DegreeCGPA_8thSem
        
        if (curr_CPI < curr_CPIThresh):
            disqualified_list[curr_COAP] = (' CPI: '+str(curr_CPI) + ' CPI Required: '+str(curr_CPIThresh))        
        if (curr_CPI >10):
            print('Invalid CPI')            
        elif curr_CPI <6.0:
            pass
            #print(str(curr_CPI) + ' '+curr_Category)
    elif curr_DegreePer_8thSem == None:
        currPerc = curr_DegreePer_7thSem
    else:
        currPerc = curr_DegreePer_8thSem        
    if currPerc >0.0: ## Valid percentage present 
        if (currPerc < currPerThresh):
            disqualified_list[curr_COAP] = (  'Percentage obtained: '+str(currPerc) + ' ,Perc Required: '+str(currPerThresh))
       
        if (currPerc >100):
            print('Invalid Percentage')            
        elif currPerc <60.0:
            pass
            #print(str(currPerc) + ' '+curr_Category)
    
    maxIdx = np.argmax(currGateScoreAll)
    maxGateScore = currGateScoreAll[maxIdx]
    maxGateRoll = currGateScoreRollNumAll[maxIdx]
    maxGateScoreDisc = currGateScoreDisc[maxIdx]
    maxGateRollDisc = maxGateRoll[:2]
    
    if maxGateScore != curr_MaxGATEScore3:
        ask_for_clarity[curr_COAP]= ('Invalid Gate Score -- Data mismatch')
        
    
    
    if maxGateRollDisc.isdigit() and maxGateScoreDisc in programCode :        
        ask_for_clarity[curr_COAP]=('Digit Gate Score '+ maxGateRoll,)
        
    elif maxGateRollDisc.upper().strip() not in programCode:
        disqualified_list[curr_COAP]= ('Invalid Gate Discipline: ' + maxGateRollDisc+ ' ,Required: '+ str(programCode))
        
    elif maxGateScoreDisc not in programCode:    
        disqualified_list[curr_COAP]= ('Invalid Gate Discipline:: ' + maxGateRollDisc+ ' ,Required: '+str(programCode))
    elif len(maxGateRoll)!=13 and len(maxGateRoll)!=11:
        ask_for_clarity[curr_COAP]=('Improper Gate Roll length '+ maxGateRoll,)
        
        
    else:
        pass
    
    maxRollWithoutDS = maxGateRoll
    maxGateRollDisc = maxGateRoll[:2]
    if maxGateRollDisc.isdigit()==False:
        maxRollWithoutDS = maxGateRoll[2:]
    row['GATE Roll num'] = maxRollWithoutDS
    currDecision, currReason = getDecision(curr_COAP, ask_for_clarity, disqualified_list)

    dfins = pd.DataFrame([['', 'Pending', '', '30/APR/2025', curr_Appno,maxRollWithoutDS,maxGateRollDisc,maxGateScore,'IIT Goa', 22, currDecision, currReason] ],columns=COLUMN_NAMES)
    dfOutput = pd.concat([dfOutput, dfins], ignore_index=True)
    if currDecision == 'Y':
        filtered_rows.append(row)

dfFiltered = pd.DataFrame(filtered_rows, columns=df.columns)

dfFiltered.to_excel(filteredData)
df.to_excel(originalData)
dfOutput.to_excel(coapData,index=False)


    
    
    
    
            