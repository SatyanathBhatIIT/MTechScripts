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

def createDataframeFromExcel(fileName, extraRows):
    wb = load_workbook(filename = fileName)
    df = pd.DataFrame(wb.worksheets[0].values)

    if extraRows == True:
        df.columns = df.iloc[2]
        df = df[3:]
    else:
        df.columns = df.iloc[0]
        df = df[1:]   
    
    column_name_list = df.columns.tolist()
    column_name_list[-1] = 'unnamed'
    if extraRows == False:
        column_name_list[-2] = 'unnamed2'
        if column_name_list[0] != 'COAP':
            column_name_list[0] = 'invalid'
    df.columns= column_name_list 
    return df


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

file1 = '/home/satyanath/data/MTech2025/CSE_raw_Mtech_2025_manual_corrections (1).xlsx'
file2 = '/home/satyanath/data/MTech2025/round0_CSE_File.xlsx'
extraRow1 = True
extraRow2 = False

df1 = createDataframeFromExcel(file1, extraRow1)
df2 = createDataframeFromExcel(file2, extraRow2)

print(df1.columns)
print(df2.columns)



primaryKeyFile1 = 'App no'
primaryKeyFile2 = 'AppNo'
#compare these cols -  Email	Contact Number	Full Name	Adm cat	Pwd	Ews	Gender	Category	COAP	GATE25RollNo	GATE25Rank	GATE25Score	GATE25Disc	GATE24RollNo	GATE24Rank	GATE24Score	GATE24Disc	GATE23RollNo	GATE23Rank	GATE23Score	GATE23Disc	MaxGATEScore out of 3 yrs	HSSC(board)	HSSC(date)	HSSC(per)	SSC(board)	SSC(date)	SSC(per)	Degree(Qualification)	Degree(PassingDate)	Degree(Branch)	Degree(OtherBranch)	Degree(Institute Name)	Degree(CGPA-7thSem)	Degree(CGPA-8thSem)	Degree(Per-7thSem)
#colsToCompare = ["Email","Contact Number","Full Name",	"Adm cat", "Pwd","Ews","Gender","Category","COAP","GATE25RollNo", "GATE25Rank","GATE25Score","GATE25Disc","GATE24RollNo","GATE24Rank","GATE24Score","GATE24Disc","GATE23RollNo","GATE23Rank","GATE23Score","GATE23Disc","MaxGATEScore out of 3 yrs","HSSC(board)","HSSC(per)","SSC(board)","SSC(date)",	"SSC(per)"	,"Degree(Qualification)"	,"Degree(PassingDate)",	"Degree(Branch)"	,"Degree(OtherBranch)"	,"Degree(Institute Name)","Degree(CGPA-7thSem)"	,"Degree(CGPA-8thSem)",	"Degree(Per-7thSem)"]
#without ['SSC(date)', 'Degree(PassingDate)']
colsToCompare = ["Email","Contact Number","Full Name",	"Adm cat", "Pwd","Ews","Gender","Category","COAP","GATE25RollNo", "GATE25Rank","GATE25Score","GATE25Disc","GATE24RollNo","GATE24Rank","GATE24Score","GATE24Disc","GATE23RollNo","GATE23Rank","GATE23Score","GATE23Disc","MaxGATEScore out of 3 yrs","HSSC(board)","HSSC(per)","SSC(board)",	"SSC(per)"	,"Degree(Qualification)",	"Degree(Branch)"	,"Degree(OtherBranch)"	,"Degree(Institute Name)","Degree(CGPA-7thSem)"	,"Degree(CGPA-8thSem)",	"Degree(Per-7thSem)"]


#create a dictionary of colname mapping: default mapping is same
colNameMapping = {}
for col in colsToCompare:
    colNameMapping[col] = col
colNameMapping['Full Name'] = 'FullName'
colNameMapping['GATE25RollNo'] = 'currYearRollNo'
#colNameMapping['GATE25Rank'] = 'currYearRank'
colNameMapping['GATE25Score'] = 'currYearScore'
#colNameMapping['GATE25Disc'] = 'currYearDisc'
colNameMapping['GATE24RollNo'] = 'prevYearRollNo'
#colNameMapping['GATE24Rank'] = 'prevYearRank'
colNameMapping['GATE24Score'] = 'prevYearScore'
#colNameMapping['GATE24Disc'] = 'prevYearDisc'
colNameMapping['GATE23RollNo'] = 'prevprevYearRollNo'
#colNameMapping['GATE23Rank'] = 'prevprevYearRank'
colNameMapping['GATE23Score'] = 'prevprevYearScore'
#colNameMapping['GATE23Disc'] = 'prevprevYearDisc'
colNameMapping['MaxGATEScore out of 3 yrs'] = 'MaxGateScore'
colNameMapping['HSSC(per)']= 'HSSCper'
colNameMapping['SSC(per)']= 'SSCper'
colNameMapping['Degree(CGPA-8thSem)']= 'DegreeCGPA8thSem'
colNameMapping['Degree(Per-8thSem)']= 'DegreePer8thSem'






presentOnlyInFile1 = {}
rowsDifferentFields = {}

#start compare
for index, row in df1.iterrows():
    coap = row[primaryKeyFile1]
    if coap not in df2[primaryKeyFile2].values:
        presentOnlyInFile1[coap] = True
    else:
        row2 = df2.loc[df2[primaryKeyFile2] == coap]
        for col in colsToCompare:
            col2 = colNameMapping[col]
            if row[col] != row2[col2].values[0]:
                if coap not in rowsDifferentFields:
                    rowsDifferentFields[coap] = []                
                rowsDifferentFields[coap].append(col)
                
#end compare
#print("Present only in file 1: ", presentOnlyInFile1.keys())
print("total present only in file 1: ", len(presentOnlyInFile1))

print("Rows with different fields: ", rowsDifferentFields)








    
    
            