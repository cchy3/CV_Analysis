#!/usr/bin/env python
# coding: utf-8

# In[104]:


import numpy as np
from math import *
import matplotlib.pyplot as plt
from matplotlib import *
import scipy
import datetime
import os
import sys
import xlrd #a library to read excel file but it won't automatically load the file with "on_demand=True" 



#-------------You have to manually cope the file name to here---------------------------------------------------------------------------------------------------------------

excelName = "CV_differen_SET_number_W18.xlsx"

#----------------------------------------------------------------------------------------------------------------------------


today_date = datetime.date.today()
currentDirectory = os.getcwd()
# check that the file path exists
if not os.path.exists(excelName):
    print('{} not found'.format(excelName))
    input("Press Enter to exit")
    sys.exit()
else:
    print("Opening the file: " + excelName + "   ... please wait")
    print("Here are the excel sheets:")
    
#use xlrd to open the excel file  
xls = xlrd.open_workbook(r'{}'.format(currentDirectory+"/"+excelName), on_demand=True)
allSheetNames = xls.sheet_names()
for i in range(len(allSheetNames)):
    print("(" + str(i) + ") " + allSheetNames[i])

#use input() to get your desired excel sheets, and check whether your input is valid
checkingInput = True
mySheetNames = []
while checkingInput:
    str_mySheetIndexes = input("-----Enter the desired sheet indexes from the list, separated by commas, do not hit SPACE-----")
    mySheetIndexes = []
    number = "" #a dummy variable
    #trying to convert string input to int array
    for char in str_mySheetIndexes:
        if char != ',':
            try:
                int(char)
                if char >= '0' and char <= '9' and char !="":
                    number +=char
                    if str_mySheetIndexes.index(char) == len(str_mySheetIndexes)-1:
                        mySheetIndexes.append(int(number))
            except:
                pass
        else:
            if len(number)!= 0:
                mySheetIndexes.append(int(number))
                number = ""
    count1 = 0
    for i in range(len(mySheetIndexes)):
        if list(enumerate(mySheetIndexes,start = 0))[i][1] in range(len(allSheetNames)):
            count1 += 1
        else:
            print(str(list(enumerate(mySheetIndexes,start = 0))[i][1]) + " is not in the list! Please re-enter your input.")
        
    if count1 == len(mySheetIndexes) and len(mySheetIndexes) !=0:   
        checkingInput = False
    elif len(mySheetIndexes) == 0 :
        print("Your input is empty, please check your input format. example input: [1,2,3]")
        
        
print("Your list of indexes is " + str(mySheetIndexes))


# In[105]:


#--------Graph Settings---------------------------------------------------------------------------------------------------------------

def getSingleIntInput(saved_var, listOfChoices, inputMessage, exceptionMessage, rangeMessage):
    checkingInput = True
    while checkingInput:
        varInput = input(inputMessage)
        try:
            int(varInput)
            if int(varInput) >= 0 and int(varInput) <= len(listOfChoices) -1:
                checkingInput = False
                saved_var = int(varInput)
                return saved_var
            else: 
                print(rangeMessage)
        except ValueError:
            print(exceptionMessage)
            pass
        
mySheets =[]
for id in mySheetIndexes:
    mySheets.append(xls.sheet_by_index(id))

sameTypeOfSheets = None
TypeOfSheets = ["Enter plot settings manually","Automatically produce the same type of plots"]
print()
print("Plot setting choices:")
for i in range(len(TypeOfSheets)):
    print("(" + str(i) + ") " + TypeOfSheets[i])
sameTypeOfSheets = getSingleIntInput(sameTypeOfSheets,TypeOfSheets,"Choose your plot setting: ", "Invalid input, please re-enter!", "Index out of range!")

mySheetNames = []
for mySheet in mySheets:
    mySheetNames.append(mySheet.name)
    
if sameTypeOfSheets == 0:
    for mySheet in mySheets:
        rw = 0
        totalRows = len(mySheet.col(0))
        for i in range(totalRows):
            if mySheet.cell(i,0).value == 'V' and mySheet.cell(i,1).value == 'I':
                break
            else:
                rw = rw+1
        varName = "" #dummy variable to pass the each variable name
        allVarNames = []
        for name_in_cell in mySheet.row(rw):
            allVarNames.append(name_in_cell.value)
        if len(allVarNames) > 4:
            print()
            print("In the sheet named " + str(mySheet.name) + "    Your analyzed CV file has these parameters: ")
            print("-----if want to choose voltage as x variable, enter index 5 instead of 0 ! -----")
        else:
            print()
            print("In the sheet named " + str(mySheet.name) + "    Your raw CV file has these parameters: ")
        for i in range(len(allVarNames)):
            print("(" + str(i) + ") " + allVarNames[i])


        xvar = None
        yvar = None
        xvar = getSingleIntInput(xvar,allVarNames,"Please choose the index of your x variable: ", "Invalid input, please re-enter!", "Index out of range!")
        yvar = getSingleIntInput(yvar,allVarNames,"Please choose the index of your y variable: ", "Invalid input, please re-enter!", "Index out of range!")

        print()

        scales = ["linear", "log"]

        x_scale = ""
        y_scale = ""
        print()
        print("Choose your scale of x and y axis: ")
        for i in range(len(scales)):
            print("(" + str(i) + ") " + scales[i])

        x_scale = getSingleIntInput(xvar,scales,"Your x axis scale will be: ", "Invalid input, please re-enter!", "Index out of range!")
        y_scale = getSingleIntInput(yvar,scales,"Your y axis scale will be: ", "Invalid input, please re-enter!", "Index out of range!")

        print("Ploting " + allVarNames[yvar] + " vs " + allVarNames[xvar] + "...")
        #start to graph
        xvarArray = []
        yvarArray = []
        for cell in range(rw+1,totalRows):
            xvarArray.append(mySheet.col(xvar)[cell].value)
            yvarArray.append(mySheet.col(yvar)[cell].value)

        get_ipython().run_line_magic('matplotlib', 'notebook')
        plt.plot(xvarArray,yvarArray,'.')
        plt.yscale(scales[y_scale])
        plt.xscale(scales[x_scale])
        plt.xlabel(allVarNames[xvar])
        plt.ylabel(allVarNames[yvar])
        pyplot.gcf().set_size_inches(12, 7, forward=True)
        plt.title(mySheet.name + "      " + allVarNames[yvar] + " vs " + allVarNames[xvar])

        if xvar == 5 and yvar ==11: #for doping vs voltage
            upperBound = max(xvarArray) - 5
            lowerBound = 20
            new_xArray = []
            new_yArray = []
            for xIndex in range(len(xvarArray)):
                if xvarArray[xIndex] >=lowerBound and  xvarArray[xIndex] <= upperBound:
                    new_xArray.append(xvarArray[xIndex])
                    new_yArray.append(yvarArray[xIndex])   
            #print(new_xArray,new_yArray)
            minDoping = min(new_yArray)
            id_minDoping = new_xArray[new_yArray.index(minDoping)]
            #print(minDoping,id_minDoping)
            #plt.plot(new_xArray,new_yArray,'.')
            plt.plot(id_minDoping,minDoping,'*',markersize = 10)
            print("Your mininum doping is " + str(minDoping) +" N/cm^3 at " + str(id_minDoping) + "V")
            print("DOPING VS VOLTAGE FINISHED")
        plt.show()
else: 
    xvar = None
    yvar = None
    scales = ["linear", "log"]

    x_scale = ""
    y_scale = ""
    rw = 0
    varName = "" 
    allVarNames = []
    
    for mySheet in mySheets:
        if mySheets.index(mySheet) == 0:
            totalRows = len(mySheet.col(0))
            for i in range(totalRows):
                if mySheet.cell(i,0).value == 'V' and mySheet.cell(i,1).value == 'I':
                    break
                else:
                    rw = rw+1
        
            for name_in_cell in mySheet.row(rw):
                allVarNames.append(name_in_cell.value)
            if len(allVarNames) > 4:
                print()
                print("In " + str(mySheetNames) + "    Your analyzed CV file has these parameters: ")
                print("-----if want to choose voltage as x variable, enter index 5 instead of 0 ! -----")
            else:
                print()
                print("In " + mySheetNames +"    Your raw CV file has these parameters: ")
            for i in range(len(allVarNames)):
                print("(" + str(i) + ") " + allVarNames[i])

            xvar = getSingleIntInput(xvar,allVarNames,"Please choose the index of your x variable: ", "Invalid input, please re-enter!", "Index out of range!")
            yvar = getSingleIntInput(yvar,allVarNames,"Please choose the index of your y variable: ", "Invalid input, please re-enter!", "Index out of range!")

            print()
            print()
            print("Choose your scale of x and y axis: ")
            for i in range(len(scales)):
                print("(" + str(i) + ") " + scales[i])

            x_scale = getSingleIntInput(xvar,scales,"Your x axis scale will be: ", "Invalid input, please re-enter!", "Index out of range!")
            y_scale = getSingleIntInput(yvar,scales,"Your y axis scale will be: ", "Invalid input, please re-enter!", "Index out of range!")

            print("Ploting " + allVarNames[yvar] + " vs " + allVarNames[xvar] + "...")
        #start to graph
        xvarArray = []
        yvarArray = []
        for cell in range(rw+1,totalRows):
            xvarArray.append(mySheet.col(xvar)[cell].value)
            yvarArray.append(mySheet.col(yvar)[cell].value)


        plt.plot(xvarArray,yvarArray,'.')
        plt.yscale(scales[y_scale])
        plt.xscale(scales[x_scale])
        plt.xlabel(allVarNames[xvar])
        plt.ylabel(allVarNames[yvar])
        pyplot.gcf().set_size_inches(12, 7, forward=True)
        plt.title(mySheet.name + "      " + allVarNames[yvar] + " vs " + allVarNames[xvar])

        if xvar == 5 and yvar ==11: #for doping vs voltage
            upperBound = max(xvarArray) - 5
            lowerBound = 20
            new_xArray = []
            new_yArray = []
            for xIndex in range(len(xvarArray)):
                if xvarArray[xIndex] >=lowerBound and  xvarArray[xIndex] <= upperBound:
                    new_xArray.append(xvarArray[xIndex])
                    new_yArray.append(yvarArray[xIndex])   
            #print(new_xArray,new_yArray)
            minDoping = min(new_yArray)
            id_minDoping = new_xArray[new_yArray.index(minDoping)]
            #print(minDoping,id_minDoping)
            #plt.plot(new_xArray,new_yArray,'.')
            plt.plot(id_minDoping,minDoping,'*',markersize = 10)
            print("Your mininum doping is " + str(minDoping) +" N/cm^3 at " + str(id_minDoping) + "V")
            print("DOPING VS VOLTAGE FINISHED")
        plt.show()


# In[ ]:


def addSheet2Excel(Excel2add,newSheetName):
    today = datetime.datetime.today()
    LastProcessed = str(today.month)+ "/" +str(today.day) + "/" + str(today.year) + "_" + str(today.hour)+":"+str(today.minute)+":"+str(today.second)
    newSheetName = xls.add_sheet(str(newSheetName))
    Excel2add.save()
    
def writeInSheet(sheet2Write,writeColNum,write):
    


# In[61]:


def WriteInNewExcel():
    print("hi")

def saveImages():
    print("hi")


# In[ ]:


#--------------will add the 1-1 correspondence of label names to the variable names in excel---------------
labelNames = ["(Negative) Bias Voltage [V]", "Current [A]", "Capacitance [F]", "Resistance [Ohm]", ""]
print(labelNames)


# In[ ]:


import xlrd, xlwt
from xlutils.copy import copy as xl_copy

# open existing workbook
existed_Excel_Name = "HPK_3.2_w11_P9_UBM_10kHz.xlsx"
existed_Excel = xlrd.open_workbook(existed_Excel_Name, formatting_info=True)
# make a copy of it
workbook = xl_copy(existed_Excel)
# add sheet to workbook with existing sheets
Sheet1 = wb.add_sheet('Sheet1')
wb.save('ex.xls')


# In[106]:


import xlrd
import os
import xlsxwriter
import math as math
import datetime
import numpy as np

#---------------------------INPUTS------------------------------------------------------------------------------------
existed_Excel_Name = "HPK_3.2_w11_P9_UBM_10kHz.xlsx"
#---------------------------------------------------------------------------------------------------------------------
today = datetime.datetime.today()
LastProcessed = str(today.month)+ "/" +str(today.day) + "/" + str(today.year) + "_" + str(today.hour)+":"+str(today.minute)+":"+str(today.second)
def getRawValuesInSheet(thisSheet): #this only works for sheet that has v, i ,c, r column
    rw = 0
    varName = "" #dummy variable to pass the each variable name
    rawVarNames = []
    totalRows = len(thisSheet.col(0))
    for i in range(totalRows):
        if thisSheet.cell(i,0).value == 'V' and thisSheet.cell(i,1).value == 'I':
            break
        else:
            rw = rw+1

    for name_in_cell in thisSheet.row(rw):
        rawVarNames.append(name_in_cell.value)
    V = [] 
    I = []
    C = []
    R = []
    for cell in range(rw+1,totalRows):
        V.append(thisSheet.col(0)[cell].value)
        I.append(thisSheet.col(1)[cell].value)
        C.append(thisSheet.col(2)[cell].value)
        R.append(thisSheet.col(3)[cell].value)
    return rawVarNames, V, I, C, R, rw
    
def generateAnalysis(excel2Analyze, areaOfSensor): 
    if not os.path.exists(excel2Analyze):
        print('{} not found'.format(excel2Analyze))
        input("Press Enter to exit")
        sys.exit()
    else:
        print(">>>Opening the file: " + excel2Analyze + "   ... please wait")
        print("Here are the excel sheets:")
    
        #use xlrd to open the excel file  
        rawExcel = xlrd.open_workbook(r'{}'.format(currentDirectory+"/"+excel2Analyze), on_demand=True)
        allSheetNames = rawExcel.sheet_names()
        for i in range(len(allSheetNames)):
            print("(" + str(i) + ") " + allSheetNames[i])
            
        if len(allSheetNames) == 1:
            print(">>>This excel file only has one sheet, trying to get raw data")
            rawSheet = rawExcel.sheet_by_index(0)
        rawData = getRawValuesInSheet(rawSheet)
        
        totalRows = len(rawSheet.col(0))
        if rawData != None:
            print("Getting raw data...done")
        else:
            print("getRawValuesInSheet returns None")
        
        Analyzed_Excel_Name = "auto-analyzed_" + existed_Excel_Name
        
        workbook = xlsxwriter.Workbook(Analyzed_Excel_Name) # create a new workbook
        newSheet = workbook.add_worksheet()
        
        print(">>>A new excel workbook has been created: " + Analyzed_Excel_Name)
        V_ = getRawValuesInSheet(rawSheet)[1]
        I_ = getRawValuesInSheet(rawSheet)[2]
        C_ = getRawValuesInSheet(rawSheet)[3]
        R_ = getRawValuesInSheet(rawSheet)[4]
        rw = getRawValuesInSheet(rawSheet)[5]
        rawDataList = [V_,I_,C_,R_]
        areaOfSensor = areaOfSensor
        #initializing more columns
        oneOverCsq_ = []
        negV_ = []
        negI_ = []
        CdivA = []
        depth = []
        AsqvCsq =[]
        derivative = []
        doping = []
        newVarList = ['V','I','C','R','1/C^2','-V','-I','C/A','Depth (um)','A^2/C^2 (cm4/F²)','Derivative (cm4/V·F²)','N (cm-3)']
        print(">>>Expanding variables..." + str(newVarList) + "...done")
        print(">>>Coping raw data...done")
        newSheet.write('H4',"area")
        newSheet.write('I4'," = ")
        newSheet.write('J4',areaOfSensor)
        try: 
            str(rawSheet.cell(7,1).value) !=None
            newSheet.write('B8',str(rawSheet.cell(7,1).value))
        except:
            print("[Warning] Cell B8 does not contain frequency of the measurement")
            pass
        #raw data passing to the new worksheet
        for colnum in range(len(newVarList)):
            newSheet.write(rw,colnum,newVarList[colnum])
            for rownum in range(rw+1,totalRows):
                if colnum <4:
                    newSheet.write(rownum,colnum,rawDataList[colnum][rownum-rw-1])
                elif colnum == 4:
                    oneOverCsq_.append(1/math.pow(C_[rownum-rw-1],2))
                    newSheet.write(rownum,colnum,1/math.pow(C_[rownum-rw-1],2)) #1/C^2
                elif colnum == 5:
                    negV_.append(np.abs(V_[rownum-rw-1]))
                    newSheet.write(rownum,colnum,np.abs(V_[rownum-rw-1])) #-V
                elif colnum == 6:
                    negI_.append(np.abs(I_[rownum-rw-1]))
                    newSheet.write(rownum,colnum,np.abs(I_[rownum-rw-1])) #-I
                elif colnum == 7:
                    CdivA.append(C_[rownum-rw-1]/areaOfSensor)
                    newSheet.write(rownum,colnum,C_[rownum-rw-1]/areaOfSensor) #C/A
                elif colnum == 8:
                    depth.append(11.9*0.0000000000000885*10000/C_[rownum-rw-1]*areaOfSensor)
                    newSheet.write(rownum,colnum,11.9*0.0000000000000885*10000/C_[rownum-rw-1]*areaOfSensor) #depth
                elif colnum == 9:
                    AsqvCsq.append(math.pow(areaOfSensor,2)/math.pow(C_[rownum-rw-1],2))
                    newSheet.write(rownum,colnum,math.pow(areaOfSensor,2)/math.pow(C_[rownum-rw-1],2)) #A^2/C^2
        '''
        for colnum in [10,11]:
            if colnum == 10:
                for rownum in range(rw+1,totalRows):
                    if rownum != totalRows-1:
                        derivative.append((AsqvCsq[rownum-(rw+1)]-AsqvCsq[rownum-(rw+1)-1])/(depth[rownum-(rw+1)]-depth[rownum-(rw+1)]))
                        newSheet.write(rownum,colnum,(AsqvCsq[rownum+1]-AsqvCsq[rownum])/(negV_[rownum+1]-negV_[rownum]))

        '''       
                
        
        print(">>>Finished analysis")   
        
        newSheet.write('H1',"Last")
        newSheet.write('I1',"Processed")
        newSheet.write('J1', LastProcessed)
        workbook.close()
        print(">>> " + str(Analyzed_Excel_Name) + " has been added to your Jupyter Notebook directory.")
        
        
currentDirectory = os.getcwd()
#must enter a float as the area(the second input), do not enter any non-numerical! 
#Exception-catching will be added later
generateAnalysis(existed_Excel_Name,0.0169) 
    


# In[89]:





# In[31]:


import xlsxwriter

workbook = xlsxwriter.Workbook('hello_world.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Hello world')

workbook.close()


# In[ ]:





# In[ ]:





# In[ ]:


#-----------------------------------------some draft codes-------------------------------------------------------------

'''
#practice: extract a number array from a string
string = "2,8,33,ab ,c,90  0, "
numbers = []
number = ""
for char in string:
    if char != ',':
        try:
            int(char)
            if char >= '0' and char <= '9' and char !="":
                number +=char
        except:
            print(char +" is not a number! Invalid input")
    else:
        if len(number)!= 0:
            numbers.append(str(number))
        number = ""
        
print(numbers)

#practice: find character in a string
arr  = range(10)
x = [1,5,7,11]

myname = "heyi"
for char in myname:
    if myname.index(char) == len(myname)-1:
        print("i")
    else:
        print("not i")

#practice: how to use enumerate
print(range(len(allSheetNames)))
numbers  = [17,18,19,20,21,30]
for i in range(len(numbers)):
    if list(enumerate(numbers,start = 0))[i][1] in range(len(allSheetNames)):
        print(str(list(enumerate(numbers,start = 0))[i][1]) + " is in the list")
    else:
        print(str(list(enumerate(numbers,start = 0))[i][1]) + " is not in the list!!!")

'''        

#----------------------------------------Abandoned codes---------------------------------------------------------------------        
'''if len(mySheetIndexes) == 0 or not str_mySheetIndexes:
            print("Your input is empty, please re-enter the indexes of the your desired sheets")
           
elif enumerate() not in range(len(allSheetNames)):
                print("You enter the wrong index: " + str(i) + ". Please only enter the number shown in the list.") 

mySheetIndexes) == 0 or 

for i in mySheetIndex:
    print(i)
mySheets = []
for id in mySheetIndexes:
    mySheet+=xls.sheet_by_index(id)
print(mysheet)'''


# In[ ]:


#some notes
#xlwt module cannot save .xlsx file, it only can save .xls file which is binary. So I change xlwt to xlsxwritter
'''
wb = xlwt.Workbook(encoding="utf-8")
ws1 = wb.add_sheet('Sheet 1',cell_overwrite_ok=True)
ws2 = wb.add_sheet('Sheet 2',cell_overwrite_ok=True)
ws3 = wb.add_sheet('Sheet 3',cell_overwrite_ok=True)
ws1.row(0).write(0, 'Data written in first cell of first sheet')

ws1.write(0, 0, 'Data overwritten in the first cell of first sheet')

ws2.write(0, 0, 'Data written in first cell of secondsheet')

ws3.write(0, 0, 'Data written in first cell of third sheet')

ws1.write(0, 1, 'Data written in first row,second column offirst sheet')

ws1.row(1).write(1, 'Data written in second row,second column of first sheet')

var = "Data from variable written in third row,second column of first sheet" 

ws1.row(2).write(1,var)
                 
wb.save('Spreadsheet_test2.xls')'''

