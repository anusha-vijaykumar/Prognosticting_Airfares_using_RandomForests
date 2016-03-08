""" Importing required Modules"""
import os
from tkFileDialog import askopenfilename
import Tkinter as tk
from Tkinter import *
import subprocess
from tkFileDialog import *
import subprocess
from datetime import datetime
import time
import ttk
from ttk import Combobox
from PIL import Image, ImageTk
import tkMessageBox
import datetime
import xlrd
from xlrd import open_workbook

#========================================================================
##TestPlanFileLocation = "banchennai (1).xlsx"
SheetName = "Sheet1"
global FlishtsLists
#print "enter1"
FlishtsLists = []
#===============================================================================
"""Creating Root(main) parent widget"""
root = Tk()
root.title('Flight Ticket Booking')
root.geometry("600x600+0+0")
canvas = Canvas(width = 100, height = 150, bg = 'red')
canvas.pack(expand = YES, fill = BOTH)

image = ImageTk.PhotoImage(file = "img9.jpg")
canvas.create_image(0, 0, image = image, anchor = NW)
image1 = ImageTk.PhotoImage(file = "img2.jpg")
canvas.create_image(0, 184, image = image1, anchor = NW)
image2 = ImageTk.PhotoImage(file = "img3.jpg")
canvas.create_image(276, 0, image = image2, anchor = NW)
image3 = ImageTk.PhotoImage(file = "img4.jpg")
canvas.create_image(276, 183, image = image3, anchor = NW)
image4 = ImageTk.PhotoImage(file = "img6.jpg")
canvas.create_image(552, 0, image = image4, anchor = NW)
image5 = ImageTk.PhotoImage(file = "img8.jpg")
canvas.create_image(460, 172, image = image5, anchor = NW)
image6 = ImageTk.PhotoImage(file = "img10.jpg")
canvas.create_image(808, 0, image = image6, anchor = NW)
image7 = ImageTk.PhotoImage(file = "img11.jpg")
canvas.create_image(719, 170, image = image7, anchor = NW)
image8 = ImageTk.PhotoImage(file = "img12.jpg")
canvas.create_image(996, 0, image = image8, anchor = NW)
image9 = ImageTk.PhotoImage(file = "img1.jpg")
canvas.create_image(1029, 195, image = image9, anchor = NW)
image10 = ImageTk.PhotoImage(file = "img5.jpg")
canvas.create_image(1255,0, image = image10, anchor = NW)
image11 = ImageTk.PhotoImage(file = "img7.jpg")
canvas.create_image(1190,123, image = image11, anchor = NW)
image12 = ImageTk.PhotoImage(file = "img6.jpg")
canvas.create_image(1190,200, image = image12, anchor = NW)

#mainloop()
#img = ImageTk.PhotoImage(Image.open("big.jpg"))
#panel = Label(root, image = img)
#panel.pack(side = "left", fill = "both", expand = "yes")
#root.mainloop()# (w, h, x, y)

#==============================================================================
"""Creating Frames....."""
mf = Frame(root)
mf.pack()
f0 = Frame(mf)
f0.pack(fill=Y)

f1 = Frame(mf)
f1.pack(fill=Y)

f2 = Frame(mf)
f2.pack(fill=Y)

f3 = Frame(mf)
f3.pack(fill=X)

f4 = Frame(mf)
f4.pack(fill=X)

f5 = Frame(mf)
f5.pack(fill=X)


#==============================================================================
""" To Disply Date.. import Datetime"""
def update_timeText():
    # Get the current time, note you can change the format as you wish

    current = time.strftime("%H:%M:%S")

    # Update the timeText Label box with the current time

    timeText.configure(text=current)

    # Call the update_timeText() function after 1 second

    root.after(1000, update_timeText)
#========================================================
def ShowChoiceForTripType():
    print v.get()
    
def ShowChoiceForFlightType():
    print va.get()

def call1(event):
    global Day
    Day = v1.get()
    print "Day = ", Day
    
def call2(event):
    global Month
    Month = v2.get()
    print "Month = ", Month
    
def call3(event):
    global Year
    Year = v3.get()
    print "Year = ", Year
    
def Fromcall(event):
    global FROM
    FROM = From.get()
    print "FROM = ", FROM
    
""" Here we are finding the row of the Selected flight in the excel
Sheet"""
def PassFlightCol(event):
    global SelectedCol                         #Selected  flights colummn
    SelectedCol = []
    SelectedFlight = v5.get()
    if SelectedFlight == "AnyOne":
        SelectedCol = DateLoc
##        SelectedColLen = len(DateLoc)
    else: 
        SelectedCol.append(SelectedFlight[0:2])
        
    print "SelectedCol = ", SelectedCol
        
def Tocall(event):
    global TO
    TO = To.get()
    print "TO = ", TO
#================================================================
""" This Function is used for confirmation after Clicking Suggest Button"""
from tkMessageBox import askyesno, showinfo  
def Suggest1():
    
    if askyesno('Verify', 'Are you Sure?'):
        root.quit()
    else:
        showinfo('Click OK', 'Now You Can Choose Again !')
    
#=================================================================
""" This funtion gives the FileLocation using input source and destination """
def GetWorkSheet():
    global TestPlanFileLocation
    if FROM == TO:
        tkMessageBox.showerror("Error", "Source and Destination should be different")
        
    elif FROM == "Bangalore" and TO == "Delhi":
        TestPlanFileLocation = "Bang2Delhi.xlsx"
        
    elif FROM == "Bangalore" and TO == "Chennai":
        TestPlanFileLocation = "Bang2Chen.xlsx"
        
    elif FROM == "Bangalore" and TO == "Mumbai":
        TestPlanFileLocation = "Bang2Mum.xlsx"
        
    elif FROM == "Delhi" and TO == "Chennai":
        TestPlanFileLocation = "Delhi2Chen.xlsx"

    elif FROM == "Delhi" and TO == "Mumbai":
        TestPlanFileLocation = "Delhi2Mum.xlsx"
        
    elif FROM == "Delhi" and TO == "Bangalore":
        TestPlanFileLocation = "Delhi2Bang.xlsx"
        
    elif FROM == "Chennai" and TO == "Bangalore":
        TestPlanFileLocation = "Chen2Bang.xlsx"
        
    elif FROM == "Chennai" and TO == "Delhi":
        TestPlanFileLocation = "Chen2Delhi.xlsx"

    elif FROM == "Chennai" and TO == "Mumbai":
        TestPlanFileLocation = "DChen2Mum.xlsx"

    elif FROM == "Mumbai" and TO == "Bangalore":
        TestPlanFileLocation = "Mum2Bang.xlsx"

    elif FROM == "Mumbai" and TO == "Delhi":
        TestPlanFileLocation = "Mum2Delhi.xlsx"

    elif FROM == "Mumbai" and TO == "Chennai":
        TestPlanFileLocation = "Mum2Chen.xlsx"

    elif FROM == "Bangalore" and TO == "Cochin":
        TestPlanFileLocation = "Bang2coch.xlsx"

    #elif FROM == "Mangalore" and TO == "Bangalore":
     #  TestPlanFileLocation = "Mang2Bang.xlsx"

   # elif FROM == "Bangalore" and TO == "Kolkata":
    #    TestPlanFileLocation = "Bang2Kolk.xlsx"

    #elif FROM == "Mumbai" and TO == "Delhi":
   #     TestPlanFileLocation = "mum2delhi.xlsx"

  #  elif FROM == "Mumbai" and TO == "Chennai":
 #       TestPlanFileLocation = "mum2chen.xlsx"

    #elif FROM == "Chennai" and TO == "Bangalore":
     #   TestPlanFileLocation = "chen2ban.xlsx"

    #elif FROM == "Chennai" and TO == "Bangalore":
     #   TestPlanFileLocation = "chen2delhi.xlsx"

    #elif FROM == "Chennai" and TO == "Mumbai":
     #   TestPlanFileLocation = "chen2mum.xlsx"

   # elif FROM == "Delhi" and TO == "Mumbai":
    #    TestPlanFileLocation = "delhi2mum.xlsx"
    

    
    return TestPlanFileLocation
            
#==============================================================================    
""" Here we are comparing Date provided by the user and the sys time, if date is
correct Then it finds the same dates in the excel sheet and later it gets Row number, flight
 name , price and departure time, after that we wil pass the details to the
 available flights Combo box. and we are calling PassFlightCol function to get
 Colummn of the selected flight and Suggest function """

def CompareTime():
    global InputDate , HeadderRows, DateLoc
    TestPlanFileLocation = GetWorkSheet()
    
    HeadderRows = 4
    now = datetime.datetime.now()
##    print "Current date and time using str method of datetime object:"
##    print str(now)
##    print "Current date and time using instance attributes:"
    years =str(now.year)
    months = str(now.month)
    days = str(now.day)
    print "Current year: %s" % years
    print "Current month: %s" % months
    print "Current day: %s" % days
    print "%s" %Day
    print Month
    print Year
    
    if (int(Day) < int(days)) or (int(Month) < int(months)):
        tkMessageBox.showerror("Error", "Wrong input")           
        
    else:
        InputDate = "a " + Day + "/" + Month + "/" + Year
        print InputDate
        Test_Workbook = xlrd.open_workbook(TestPlanFileLocation)
        Test_Worksheet = Test_Workbook.sheet_by_name(SheetName)
    
        Test_Nrows = Test_Worksheet.nrows
        
        print 'Test_Nrows = ',Test_Nrows
        Test_Ncols = Test_Worksheet.ncols
        DateCount = 0
        DateLoc = []
        
        
        for i in range(HeadderRows,Test_Nrows):
            #print "enter1"
            
            Date = Test_Worksheet.cell_value(i,5)
            print "date = ",Date
            if Date == InputDate:
                print "match"
                DateCount += 1
                DateLoc.append(i)
        print "DateLoc = ", DateLoc
        print "DateCount = ", DateCount
        for dateCount in range (0,DateCount):
            FL = str(Test_Worksheet.cell_value(DateLoc[dateCount],19))# to get selected flight
            PL = str(Test_Worksheet.cell_value(DateLoc[dateCount],10)) # to get selected flight's price
            DEPL = str(Test_Worksheet.cell_value(DateLoc[dateCount],7)) # to get selected flight's departure time
            flightColumn = str(DateLoc[dateCount]) # to get corresponding column of the selected flight
            print "enter"
            ADDL = flightColumn + "    Flight = " + FL + "    Price = " + PL + "Rs" + "    Dept Time = " + DEPL
            FlishtsLists.append(ADDL) # print's when available flight's button is pressed
            print "exit"
##            FlishtsListsCol.append()
        FlishtsLists.append("AnyOne")  # adds anyone option too
            
            
        print "FlishtsLists = ",FlishtsLists
        #=============================================================================
        """Available flight Label"""

        Label(f4, text="Available Flights", fg = "red",font=('Lucida Sans Typewriter',(15), "bold")).pack(side = LEFT,
                                                                             padx = 10,
                                                                             pady = 40)

##    f4 = Frame(mf)
##    f4.pack(fill=X)
        global v5 
        v5 = StringVar() # will take selected flights value
        Values = FlishtsLists
        w = Combobox(f4,values = Values ,width = 60, height = 5, textvariable= v5)
        w.bind('<<ComboboxSelected>>',PassFlightCol)
        w.pack(side = LEFT,padx = 10, pady = 40)

        Button(f4, text="SUGGEST",fg ="white" ,bg = "blue" ,font=('Lucida Sans Typewriter',(10), 'bold'),
           command=lambda: Suggest1()).pack(side = LEFT, padx=10,pady=40)
    return


#==============================================================================
""" TO create Flight Label, here bg= baground colour, fg = text color"""

FlightLabel = """ Aviator GURU"""
msg = Message(f0, text = FlightLabel)
msg.config(bg='navy blue', font=('Book Antiqua', 40, 'italic', 'bold'),width = 1000, relief = RAISED,
           borderwidth = 5, fg = "white")
msg.pack(side = "left", pady =30 )


#===========================================================================
""" From Label"""

Label(f1, text="From",bg="white", fg = "navy blue", anchor="w", font=('Lucida Sans Typewriter',(15))).pack(side = LEFT,
                                                                     padx = 10,
                                                                     pady = 20)
""" Como box for "From/ source" """

From = StringVar()
Values = ["Bangalore", "Delhi", "Mumbai", "Chennai"]
w = Combobox(f1,values = Values,width = 15, height = 4, textvariable= From)
w.bind('<<ComboboxSelected>>', Fromcall)
w.pack(side = LEFT, padx = 10, pady = 20)

""" TO Label"""

Label(f1, text="To",bg="white",  fg = "navy blue",font=('Lucida Sans Typewriter',(15))).pack(side = LEFT,padx = 10, pady = 20)

""" Como box for "To/Destination" """
                                                                         
To = StringVar()
Values = ["Bangalore", "Delhi", "Mumbai", "Chennai", "Kolkata", "Cochin"]
w = Combobox(f1,values = Values,width = 15, height = 4, textvariable= To)
w.bind('<<ComboboxSelected>>', Tocall)
w.pack(side = LEFT, padx = 10, pady = 20)

#===========================================================================

""" Departure Time Label"""                  

L1 = Label(f2, text="Departure Date  ",bg="white", fg = "navy blue",font=('Lucida Sans Typewriter',(15)))
L1.grid(row=0, column=1, padx= 20,pady= 40,sticky="e")
L1.pack(side = LEFT)
#=================================
""" DD Label and Combo box"""

Label(f2, text="DD", fg = "blue", font=('Berlin Sans FB Demi',(15))).pack(side = LEFT,
                                                                         pady = 40)
v1 = StringVar()
Values = []
for ik in range(32):
    Values.append(ik)
w = Combobox(f2,values = Values,width = 10, height = 5, textvariable= v1)
w.bind('<<ComboboxSelected>>', call1)
w.pack(side = LEFT, padx = 10, pady = 40)
#======================================

""" MM Label and Combo box"""

Label(f2, text="MM", fg = "blue",font=('Berlin Sans FB Demi',(15))).pack(side = LEFT, pady = 40)

v2 = StringVar()
Values = []
for ik in range(13):
    Values.append(ik)
w = Combobox(f2,values = Values,width = 10, height = 5, textvariable= v2)
w.bind('<<ComboboxSelected>>', call2)
w.pack(side = LEFT, padx = 10, pady = 40)
#=============================================

""" YY Label and Combo box"""
Label(f2, text="YY", fg = "blue",font=('Berlin Sans FB Demi',(15))).pack(side = LEFT, pady = 40)

v3 = StringVar()
Values = []
for ik in range(5):
    Values.append(2015+ik)
w = Combobox(f2,values = Values,width = 10, height = 5, textvariable= v3)
w.bind('<<ComboboxSelected>>', call3)
w.pack(side = LEFT, padx = 10, pady = 40)


#==============================================================================
##root.wm_title("Simple Clock Example")

# Create a timeText Label (a text box)

timeText = tk.Label(root, text="", font=("Lucida Sans Typewriter", 10))

timeText.pack(side = RIGHT)

update_timeText()
#=============================================================================
""" Search Button """

Button(f3, text="Search Flight",fg ="blue" ,bg = "white" ,font=('Lucida Sans Typewriter',(10)), width=20,
       command=lambda: CompareTime()).pack(padx=80,pady=10)


#==============================================================================
root.mainloop()
#==============================================================================
""" Tree Creating Script"""


import os
import xlrd
from xlrd import open_workbook
from xlwt import Workbook, easyxf
from xlutils.copy import copy
import csv
import math
import datetime
#==============================================================================
WBookName = TestPlanFileLocation
SheetName = "Sheet1"
ClassCol = 23 # Class Colummn( Buy/Wait )
rating = 0 
#==============================================================================
""" this funtion exits from the Result window"""
def Quit(root1):
    root1.quit()
#==============================================================================
""" output window (Message Box)"""
def DisplayMessage(Msg):
    root1 = Tk()
    mf1 = Frame(root1)
    mf1.pack()
    """ Frame """
    Fr1 = Frame(mf1)
    Fr1.pack(fill=Y)
    root1.title('RESULT') # to create display window
    root1.geometry("350x210+800+200") # (w, h, x, y)
    msg = Message(root1, text = Msg)
    msg.config(bg='lightgreen', font=('Lucida Sans Typewriter', 14, 'italic'),width = 400, relief = GROOVE,
               borderwidth = 10)
    msg.pack( )
    #=============================================================================
    """ OK Button """

    Button(Fr1, text="OK",fg ="black" ,bg = "navy blue" ,font=('Lucida Sans Typewriter',(10)), width=10,
       command=lambda: Quit(root1)).pack(padx=30,pady=10)

    mainloop( )
#=============================================================================================
""" Compares lists"""
 
def comp(list1, list2):                  # for info gain calculation
    if list1[0] == list2[0]:
        if list1[1] == list2[1]:
             if list1[2] == list2[2]:
                 return "equal"
    else:
        return "not equal"
            


#==============================================================================
def openwbook(TestPlanFileLocation):
    print "Opening Workbook............"
    
    Test_Workbook = xlrd.open_workbook(TestPlanFileLocation)
      
    return  Test_Workbook


def RowCount( Test_Workbook, SheetName):
    global Test_Nrows, DateCount
    global Test_Worksheet
    Test_Worksheet = Test_Workbook.sheet_by_name(SheetName)
    
    Test_Nrows = Test_Worksheet.nrows
    
    print 'Test_Nrows = ',Test_Nrows
    Test_Ncols = Test_Worksheet.ncols
    print 'Test_Ncols = ',Test_Ncols
    return Test_Worksheet, Test_Nrows, Test_Ncols

def log2( x ):
    if x != 0.0:
        return math.log( x ) / math.log( 2 )
    else: return (0)
    
def ClassEntropy(ClassCol,rating):
    global DateCount, DateLoc

    HeadderRows = 4
    Test_Workbook = openwbook(WBookName)
    Test_Worksheet, Test_Nrows, Test_Ncols = RowCount(Test_Workbook, SheetName)
    TotalSamples = Test_Nrows - HeadderRows
    print "TotalSamples", TotalSamples
    
    Class1 = 0
    Class2 = 0
    Class3 = 0
    
    DateCount = 0
    DateLoc = []
    for i in range(HeadderRows,Test_Nrows):
        #=====================================================
               
        Date = Test_Worksheet.cell_value(i,5)
##        print "date = ",Date
        if Date == InputDate:
            print "test"
            DateCount += 1
            DateLoc.append(i)      

        #======================================================   
        TestClass = Test_Worksheet.cell_value(i,ClassCol)
##        print TestClass
        if rating == 0:
            if TestClass == 1:
                Class1 += 1.0
            elif TestClass == 2:
                Class2 += 1.0
            elif TestClass == 3:
                Class3 += 1.0
        else:
            if TestClass == 4:
                Class1 += 1.0
            elif TestClass == 6:
                Class2 += 1.0
            elif TestClass == 10:
                Class3 += 1.0
## ==============================================================              
    print "DateLoc = ", DateLoc
    print "DateCount = ", DateCount
    for dateCount in range (0,DateCount):
        print Test_Worksheet.cell_value(DateLoc[dateCount],19)
#================================================================    
    print "Class1 = ", Class1
    print "Class2 = ", Class2
    print "Class3 = ", Class3
    """ Calculating entropy """
    A = float(Class1/TotalSamples)
    B = float(Class2/TotalSamples)
    C = float(Class3/TotalSamples)
    print A
    LogA = log2(A)
    LogB = log2(B)
    LogC = log2(C)
    print "LogA = ",LogA
    InfoT = (-(A * LogA))+(-(B * LogB))+(-(C * LogC))
    ##InfoT = -(A * LogA2)
    print"InfoT = ", InfoT
    return InfoT, TotalSamples,Test_Worksheet, Test_Nrows


def AttributesEntropy(TrainData,ClassCol,AttributeCol,rating):
    global HeadderRows
    HeadderRows = 4
    Attri012 = 0.0
    AttriC1 = 0
    AttriC2 = 0
    AttriC3 = 0
    
    
    for j in range(HeadderRows,Test_Nrows):
        TestStop = Test_Worksheet.cell_value(j,AttributeCol)
        TestClass = Test_Worksheet.cell_value(j,ClassCol)
    ##    print "TestStop", TestStop
     ##    print "TestClass", TestClass
        if rating == 0:
            if TestStop == TrainData and TestClass == 1:
                Attri012 += 1.0
                AttriC1 += 1.0
            elif TestStop == TrainData and TestClass == 2:
                Attri012 += 1.0
                AttriC2 += 1.0
            elif TestStop == TrainData and TestClass == 3:
                Attri012 += 1.0
                AttriC3 += 1.0
        else:
            if TestStop == TrainData and TestClass == 4:
                Attri012 += 1.0
                AttriC1 += 1.0
            elif TestStop == TrainData and TestClass == 6:
                Attri012 += 1.0
                AttriC2 += 1.0
            elif TestStop == TrainData and TestClass == 10:
                Attri012 += 1.0
                AttriC3 += 1.0
            
    return Attri012,AttriC1,AttriC2,AttriC3
#==============================================================================  

attribute = 0
InfoT, TotalSamples,Test_Worksheet, Test_Nrows = ClassEntropy(ClassCol,rating)
def GainCal(rating, ClassCol,AttributeCol):
    


    TrainData = 1
    Attri1OfZeroes, Attri1OfZeC1 ,Attri1OfZeC2 ,Attri1OfZeC3 = AttributesEntropy(TrainData,
                                                                           ClassCol,
                                                                           AttributeCol,
                                                                           rating)
    TrainData = 2
    Attri1OfOnes, Attri1Of1C1 , Attri1Of1C2 ,Attri1Of1C3 = AttributesEntropy(TrainData,
                                                                           ClassCol,
                                                                           AttributeCol,
                                                                           rating)

    TrainData = 3
    Attri1OfTwo, Attri1Of2C1 ,Attri1Of2C2 ,Attri1Of2C3 = AttributesEntropy(TrainData,
                                                                           ClassCol,
                                                                           AttributeCol,
                                                                           rating)

##    print "Attri1OfZeroes = ", Attri1OfZeroes
##    print "Attri1OfOnes = ", Attri1OfOnes
##    print "Attri1OfTwo = ", Attri1OfTwo
##    print "Attri1OfZeC1 = ",Attri1OfZeC1
##    print "Attri1OfZeC2 = ",Attri1OfZeC2
##    print "Attri1OfZeC3 = ",Attri1OfZeC3
##    print "Attri1Of1C1 = ",Attri1Of1C1
##    print "Attri1Of1C2 = ",Attri1Of1C2
##    print "Attri1Of1C3 = ",Attri1Of1C3
##    print "Attri1Of2C1 = ",Attri1Of2C1
##    print "Attri1Of2C2 = ",Attri1Of2C2
##    print "Attri1Of2C3 = ",Attri1Of2C3
##
    """ Calculating entropy """




    if Attri1OfZeroes !=0:
        AZero = float(Attri1OfZeC1/Attri1OfZeroes)
        BZero = float(Attri1OfZeC2/Attri1OfZeroes)
        CZero = float(Attri1OfZeC3/Attri1OfZeroes)
        
        LogAZero = log2(AZero)
        LogBZero = log2(BZero)
        LogCZero = log2(CZero)
        StopZeros = float(Attri1OfZeroes/TotalSamples)

        InfoT0 = StopZeros*((-(AZero * LogAZero))+(-(BZero * LogBZero))+(-(CZero * LogCZero)))

    if Attri1OfOnes !=0:
        AOne = float(Attri1Of1C1/Attri1OfOnes)
        BOne = float(Attri1Of1C2/Attri1OfOnes)
        COne = float(Attri1Of1C3/Attri1OfOnes)
        
        LogAOne= log2(AOne)
        LogBone = log2(BOne)
        LogCOne = log2(COne)
        StopOnes = float(Attri1OfOnes/TotalSamples)
        InfoT1 = StopOnes*((-(AOne * LogAOne))+(-(BOne * LogBone))+(-(COne * LogCOne)))
##        print InfoT1

        

    if Attri1OfTwo != 0:
        ATwo = float(Attri1Of2C1/Attri1OfTwo)
        BTwo = float(Attri1Of2C2/Attri1OfTwo)
        CTwo = float(Attri1Of2C3/Attri1OfTwo)

        LogATwo= log2(ATwo)
        LogBTwo = log2(BTwo)
        LogCTwo = log2(CTwo)
        StopTwo = float(Attri1OfTwo/TotalSamples)
        InfoT2 = StopTwo*((-(ATwo * LogATwo))+(-(BTwo * LogBTwo))+(-(CTwo * LogCTwo)))


    if Attri1OfTwo != 0:    
        InfoTX3 = InfoT0 + InfoT1 + InfoT2
    elif Attri1OfOnes !=0:
        InfoTX3 = InfoT0 + InfoT1
    elif Attri1OfZeroes !=0:
        InfoTX3 = InfoT0 
        
        

##    print"InfoT = ", InfoTX3
    Gain = InfoT - InfoTX3
##    print "Gain = ", Gain
    return Gain
#===================================================================================

rating = 0
SortGainList , EqualOnes, Buy ,Wait, UpToYou = [], 0, 0, 0, 0
Icount = 0
def Tree(ColList,SortGainListInput):
    global GainX1,GainX2,GainX3
    print "Tree================================="
    AttributeCol1 = ColList[0]    # departure time colummn
    GainX1 = GainCal(rating, ClassCol,AttributeCol1)
    AttributeCol2 = ColList[1]   # Stops colummn
    ##TestClass = Test_Worksheet.cell_value(j,ClassCol)
    GainX2 = GainCal(rating, ClassCol,AttributeCol2)
    AttributeCol3 = ColList[2]   # Airline colummn
    GainX3 = GainCal(rating, ClassCol,AttributeCol3)
    SortGain = [[GainX1,AttributeCol1],[GainX2,AttributeCol2],[GainX3,AttributeCol3]]
    SortGain.sort(reverse=True)
    SortGainList = SortGainListInput
    EqualOnes ,Icount,Buy ,Wait, UpToYou =0, 0 ,0,0,0
    ##            print "SortGainList = ", SortGainList
    for TestRow in range (HeadderRows,Test_Nrows):
        TrainDataList = []
        
        for non, AttributeCol in SortGain:
            Attri1 = Test_Worksheet.cell_value(TestRow,AttributeCol)
            TrainDataList.append(Attri1)
##            print "TrainDataList = ",TrainDataList
        ans = comp(SortGainList, TrainDataList)
##        print "ans = ", ans
        if ans == "equal":
            EqualOnes += 1
##            print "TrainDataList = ",TrainDataList
            Temp = Test_Worksheet.cell_value(TestRow,ClassCol)
            if Temp == 1:
                Buy += 1
            elif Temp == 2:
                Wait += 1
            elif Temp == 3:
                UpToYou += 1

    Icount += 1
##    print EqualOnes
##    print "Buy =", Buy
##    print "Wait = ", Wait
##    print "UpToYou = ", UpToYou
##    print "Icount = ",Icount


    if (Buy > Wait) and (Buy > UpToYou) and (Buy != Wait) and (Buy != UpToYou) :
##        print "You can BUY..........:-)"
        return "Buy"
    elif (Wait > Buy) and (Wait > UpToYou) and (Wait != Buy) and (UpToYou != Wait):
##        print "you better Wait............."
        return "Wait"
    elif (UpToYou > Buy) and (UpToYou > Wait) and (UpToYou != Buy) and (UpToYou != Wait):
##        print "UpToYou............."
        return "UpToYou"
    elif Buy == Wait == UpToYou == 0 :
##        print "Invalid............."
        return "Invalid"
    elif (Buy == Wait) or (Wait ==UpToYou) or (Buy == UpToYou):
##        print "Can't Decide"
        return "Can't Decide"


#==============================================================================
def GainListInput(AttributeCol,daterow):
    SortGainListInput = []
    global GainX1,GainX2,GainX3
##    print "Tree1================================="
    AttributeCol1 = AttributeCol[0]    # departure time colummn
    GainX1 = GainCal(rating, ClassCol,AttributeCol1)
    AttributeCol2 = AttributeCol[1]   # Stops colummn
    ##TestClass = Test_Worksheet.cell_value(j,ClassCol)
    GainX2 = GainCal(rating, ClassCol,AttributeCol2)
    AttributeCol3 = AttributeCol[2]   # Airline colummn
    GainX3 = GainCal(rating, ClassCol,AttributeCol3)
    SortGain = [[GainX1,AttributeCol1],[GainX2,AttributeCol2],[GainX3,AttributeCol3]]
##    for dateCount in range (0,DateCount):
##        Row = DateLoc[dateCount]
    SortGain.sort(reverse=True)
##    print "SortGain[0][1] = ", SortGain[0][1]
##    print "SortGain[0][1] = ", SortGain[1][1]
##    print "SortGain[0][1] = ", SortGain[2][1]
##    

    Attri1 = Test_Worksheet.cell_value(daterow,SortGain[0][1])
    SortGainListInput.insert(0,Attri1)
    Attri1 = Test_Worksheet.cell_value(daterow,SortGain[1][1])
    SortGainListInput.insert(1,Attri1)
    Attri1 = Test_Worksheet.cell_value(daterow,SortGain[2][1])
    SortGainListInput.insert(2,Attri1)
##    print "SortGainListInput = ", SortGainListInput
    return SortGainListInput
#==============================================================================

def GetbestResult(FinalResult):
    Buy,Wait, UpToYou, Invalid, CantDecide = 0, 0, 0, 0, 0
    for kl in range (0,len(FinalResult)):
        if FinalResult[kl] == "Buy":
            Buy += 1
        elif FinalResult[kl] == "Wait":
            Wait += 1
        elif FinalResult[kl] == "UpToYou":
            UpToYou += 1
        elif FinalResult[kl] == "Invalid":
            Invalid += 1
        elif FinalResult[kl] == "Can't Decide":
            CantDecide += 1
    if (Buy > Wait) and (Buy > UpToYou) and (Buy != Wait) and (Buy != UpToYou) :
##        print "You can BUY..........:-)"
        return "Buy"
    elif (Wait > Buy) and (Wait > UpToYou) and (Wait != Buy) and (UpToYou != Wait):
##        print "you better Wait............."
        return "Wait"
    elif (UpToYou > Buy) and (UpToYou > Wait) and (UpToYou != Buy) and (UpToYou != Wait):
##        print "UpToYou............."
        return "UpToYou"
    else:
##        print "Can't Decide"
        return "UpToYou"

#==================================================================================

    
"""Creating Trees... here u acn give as many tree with max of 3 attributes"""

tree1 = [8,18,20]
tree2 = [20,14,16]
tree3 = [11,14,8]
tree4 = [11,14,16]
tree5 = [18,14,8]


HeadderRows = 4
FinalDecision = []

for DR in range (0,len(SelectedCol)):
    print "len(SelectedCol) = ",len(SelectedCol)
    Date = SelectedCol[DR]
    DateRows = int(Date)
    print "Date = ", Date
    print "DateRows = ", DateRows
    FinalResult = []
    SortGainListInput = []
    ##    DateRow = 1
    
    AColList = [tree1,tree2,tree3,tree4 ,tree5]
    for s in range(0,5):
        
        AttributeCol = AColList[s]
        gainValue = GainListInput(AttributeCol,DateRows)
        SortGainListInput.append(gainValue)
    ##        print "AColList[s] = ",AColList[s]
    ##        print "SortGainListInput[s] = ", SortGainListInput[s]
        Result = Tree(AColList[s],SortGainListInput[s])
    ##        print "Result = ", Result
        FinalResult.append(Result)
    ##    print "FinalResultList = ", FinalResult
    Decision = GetbestResult(FinalResult)
    print "final result to the user is ", Decision
    FinalDecision.append(Decision)
    print "FinalDecision = ", FinalDecision
    
    
if len(SelectedCol) > 1:
    MsgToBeDisplayed = ""
    for Fly in range (len(SelectedCol)):
        FlightName = Test_Worksheet.cell_value(SelectedCol[Fly],19)
        print "FlightName = ",FlightName
        MsgToBeDisplayed = MsgToBeDisplayed +"Flight = %s,   Suggetion is %s" %(FlightName ,FinalDecision[Fly]) + "\n"
    print MsgToBeDisplayed
elif Decision == "Buy":
    MsgToBeDisplayed = """chosen Flight with specfied date
for the required journey is avilable.
    \n We Suggest You To %s""" % Decision
elif Decision == "Wait":
    MsgToBeDisplayed = """chosen Flight with specfied date
is not satisified your Requirements.
    \n We Suggest You To %s""" % Decision

elif Decision == "UpToYou":
    MsgToBeDisplayed = """chosen Flight with specfied date
for the required journey is avilable.\n We Suggest You To %s""" % Decision

DisplayMessage(MsgToBeDisplayed) # to output the result

subprocess.Popen(r'python GenericFlightContinuTreeNext.py')# Upon Clickin ok Button
# on the Output/ result window the scripts exits/quits from execution in order to take
#get/ put next input we will call the script/ run the script using Subprocess command

    

        


    
    
    


        
    

    
            






















