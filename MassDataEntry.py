#!/usr/bin/env python3
import os
import docx
import pandas as pd
import openpyxl
import datetime

def prepare_xldata(Excel_dataframe):
    dict={}
    header_list=Excel_dataframe.columns.ravel()  # List of column headers.
    for headerName in header_list:
        dict[headerName]=[]
        dict[headerName]=Excel_dataframe["{}".format(headerName)].tolist() #list of all values in ccorresponding column
    
    # We take the needed columns from the dictionary and save as list.
    employeeName = dict['EMPLOYEE NAME']
    passport = dict['NRIC/PASSPORT NO.']
    jobtitile = dict['JOB TITLE']
    employmentDate = dict['EMPLOYEMENT DATE']
    employmentType = dict['EMPLOYEMENT TYPE']
    workingHours = dict['WORKING HOURS PER WEEK']
    bonus = dict['INCENTIVES*']
    salary = dict['SALARY']
    # To fetch from different cell above the current header.
    tempDf = pd.read_excel(excelFile, header=0) # change header value to access the "Branch name" cell.
    branch = tempDf.iloc[1,2] # branch name is in [1,2] cell from 0 header.
    for i in range(len(employeeName)):
         format_in_words(employeeName[i],passport[i],jobtitile[i],employmentDate[i],salary[i],employmentType[i],workingHours[i],bonus[i],branch)
    
def format_in_words(employeeName,passport,jobtitile,employmentDate,salary,employmentType,workingHours,bonus,branch):
        
    doc = docx.Document("sample letter of verification.docx")
    all_para= doc.paragraphs
    incentive=str(bonus)
    single_para=all_para[15]
    if incentive != "nan":
        incentive="Plus"+" "+str(incentive)
        single_para.text=single_para.text.replace("[include bonus frequency, if applicable]","{}.".format(incentive))
    else:
        incentive=" " 
        single_para.text=single_para.text.replace("[include bonus frequency, if applicable]","{}".format(incentive))    
    
    for para in all_para:
        # Format the existing doc according to our template.
        para.text=para.text.replace("[Employee name]","{}".format(employeeName))
        para.text=para.text.replace("[NRIC/Passport no.]","{}".format(passport))
        para.text=para.text.replace("[Branch name]","{}".format(branch))
        para.text=para.text.replace("[employee job title]","{}".format(jobtitile))
        para.text=para.text.replace("[employment date]","{}".format(employmentDate))
        para.text=para.text.replace("[full-time or part-time]","{}".format(employmentType))
        para.text=para.text.replace("[number]","{}".format(workingHours))
        para.text=para.text.replace("[amount]","{}".format(salary))
        
        para.text=para.text.replace("[Date]","{}".format(datetime.datetime.today().strftime('%d/%m/%Y')))
    
    doc.save("{}/{}.docx".format(os.getcwd(),employeeName))    
            
excelFile= "staff detail.xlsx"
Excel_dataframe = pd.read_excel(excelFile, header=5) # For this sample, head start at row(6)-1
Excel_dataframe = Excel_dataframe[Excel_dataframe.filter(regex='^(?!Unnamed)').columns] # Removes Unnamed columns
prepare_xldata(Excel_dataframe)