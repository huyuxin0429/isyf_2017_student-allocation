"""
ISYF 2017 student allocation program
Takes in list of student names and 2 preferences
Ouputs in Results.xlsx

By:
Hu Yu Xin

"""
import csv
import openpyxl
import os
import random
from openpyxl import Workbook,load_workbook
#Initialisation

#Output initialisation
wboutput = Workbook()   #Creating output workbook
wsoutput = wboutput.active #Creating output worksheet

wsoutput["A1"] = "Student"   #Creating output headers
wsoutput["B1"] = "Class Allocation"
wsoutput["C1"] = "Dialogue Allocation"
#Output initialisation end

#Classes array initialisation
classes =[]                 #Definiting array for class allocation
for x in range(0,4):
    classes.append([0])     #Number of student in each class
    
classes[0].append(48)       #Second number determines maximum class size
classes[1].append(24)
classes[2].append(24)
classes[3].append(24)
#Classes array initialisation end

dialogue = []
for x in range(0,5):
    dialogue.append([0])
dialogue[0].append(80)
dialogue[1].append(80)
dialogue[2].append(80)
dialogue[3].append(80)
dialogue[4].append(80)

debug = bool(int(input("Debug view? Enter 1 for yes, 0 for no ")))           #Debug view toggle
#debug = True

allocationIndex = ["Bio1","Bio2","Phy","Comp"] #Index for each class
"""Reference index for classes
0: Bio1
1: Bio2
2: Phy
3: Comp
"""
dialogueIndex = ["Math","Com","Phy","Bio","Chem"] #Index for each class
"""Reference index for classes
0: Math
1: Com
2: Phy
3: Bio
4: Chem
"""
# Input initialisation
filename = input("Name of source file? Write in full Eg. Student_Choice.xlsx ") #User input source file name
file_path = os.path.join(os.getcwd()+"/"+filename)
inputworkbook = load_workbook(file_path)
inputworksheet = inputworkbook.active

customise = bool(int(input(
"""Customise input data range?
Default options:
Header(Title rows, no data): 1 row
Name colomn: 1
Prefernce 1 colomm: 2
Preference 2 colomn: 3
Enter 1 for yes, 0 for no """
)))
#customise = False
if customise:
    header = int(input("Enter header rows:"))
    namecol =int(input("Enter name colomn no.:"))
    choice1col =int(input("Enter choice 1 colomn no.:"))
    choice2col =int(input("Enter choice 2 colomn no.:"))
else:
    header = 1
    namecol = 1
    choice1col = 2
    choice2col = 3
# Input initialisation end

#End Initialisation

def stucount(worksheet): #Counting and returning total no. of students by finding no. of lines to first null value
    x = 2 # row start at 2 to accomodate for first row headers
    while worksheet.cell(row = x, column = 1).value != None: 
        x += 1
    print("Total no. of students:",x)
    return x-2
"""
Preference Choices:
0: math
1: comp
2: phy
3: bio
4: chem
"""
preferenceIndex=["math","comp","phy","bio","chem"]

def allocate(preference1,preference2):
    #Allocation selector, insert 2 preference, output allocation
    #Allocate by preference and class availibility
    
    preference1 -= 1 #Convert from range 1-5 to 0-4
    preference2 -= 1
    if debug:
        print("Student choice:",preferenceIndex[preference1],preferenceIndex[preference2])
    choicetoclass = {}
    choicetoclass[0] = 2 #All math choices are equivalent to phy choices
    choicetoclass[1] = 3
    choicetoclass[2] = 2
    choicetoclass[3] = 0
    choicetoclass[4] = 0 #All Chem choices are equivalent to bio choices
    allocated = False
    preference=[preference1,preference2]
    while allocated == False: #Allocation status flag
       
        computedclass=[0,0]
        for u in range(0,2):
            computedclass[u] = choicetoclass[preference[u]]
        prefincre = 0 
        while prefincre <=1:
            if debug:
                    if prefincre == 0:
                        print("Allocation by preference 1...",preference1,preferenceIndex[preference1])
                    else:
                        print("Allocation by preference 2...",preference2,preferenceIndex[preference2])
            if computedclass[prefincre] == 0:#Bio choice have 2 class and get to be allocate
                if debug:
                    print("Bio choice")
                if classes[1][0]<classes[1][1]: 
                    classes[1][0] +=1
                    if debug:
                        print("Bio2 class size:",classes[1][0])
                    allocated = True
                    return 1
                elif classes[0][0]<classes[0][1]:
                    classes[0][0] +=1
                    if debug:
                        print("Bio1 class size:",classes[0][0])
                    allocated = True
                    return 0
            elif classes[computedclass[prefincre]][0]<classes[computedclass[prefincre]][1]:
                if debug:
                    print("Not bio choice")
                classes[computedclass[prefincre]][0] +=1
                if debug:
                        print(allocationIndex[computedclass[prefincre]],"class size:",classes[computedclass[prefincre]][0])
                allocated = True 
                return computedclass[prefincre]
            else:
                prefincre += 1
                if debug:
                    if prefincre == 1:
                        print("Prefernce 1 allocation failed")
                    else:
                        print("Prefernce 2 allocation failed")
        while allocated == False:   #If both preferences are unavilible, randomly selects a class
            if debug:
                print("Assigning to random availible class...")
            classarray = [0,1,2,3]
            for x in range(0,2):
                if computedclass[x] == 0:
                    classarray.remove(0)
                    classarray.remove(1)#Removing preferences from availible classes
                else:
                    classarray.remove(computedclass[x])
            randomclass = random.randint(0,len(classarray)-1)
            if debug:
                print(randomclass)
            allocatedclass = classarray[randomclass]    #Randomly selecting among last  class
            if classes[allocatedclass][0]<classes[allocatedclass][1]:
                classes[allocatedclass][0]+=1
                if debug:
                        print(allocationIndex[allocatedclass],"class size:",classes[allocatedclass][0])
                allocated = True
                return allocatedclass
def allocateDialogue(preference1,preference2):
    #Allocation selector, insert 2 preference, output allocation
    #Allocate by preference and lecture availibility
    
    preference1 -= 1 #Convert from range 1-5 to 0-4
    preference2 -= 1
    if debug:
        print("Student choice:",preferenceIndex[preference1],preferenceIndex[preference2])
    allocated = False
    preference=[preference1,preference2]
    while allocated == False: #Allocation status flag
        computeddialogue=[0,0]
        for u in range(0,2):
            computeddialogue[u] = preference[u]
        prefincre = 0 
        while prefincre <=1:
            if debug:
                    if prefincre == 0:
                        print("Allocation by preference 1...",preference1,preferenceIndex[preference1])
                    else:
                        print("Allocation by preference 2...",preference2,preferenceIndex[preference2])
            if dialogue[computeddialogue[prefincre]][0]<dialogue[computeddialogue[prefincre]][1]:
                dialogue[computeddialogue[prefincre]][0] +=1
                if debug:
                        print(allocationIndex[computeddialogue[prefincre]],"class size:",dialogue[computeddialogue[prefincre]][0])
                allocated = True 
                return computeddialogue[prefincre]
            else:
                prefincre += 1
                if debug:
                    if prefincre == 1:
                        print("Prefernce 1 allocation failed")
                    else:
                        print("Prefernce 2 allocation failed")
        while allocated == False:   #If both preferences are unavilible, randomly selects a class
            if debug:
                print("Assigning to random availible class...")
            dialoguearray = [0,1,2,3,4]
            for x in range(0,2):
                if computeddialogue[x] == 0:
                    dialoguearray.remove(0)
                    dialoguearray.remove(1)#Removing preferences from availible classes
                else:
                    classarray.remove(computedclass[x])
            randomdialogue = random.randint(0,len(classarray)-1)
            if debug:
                print(randomdialogue)
            allocateddialogue = dialoguearray[randomdialogue]    #Randomly selecting among last  class
            if dialogue[allocateddialogue][0]<dialogue[allocateddialogue][1]:
                dialogue[allocateddialogue][0]+=1
                if debug:
                        print(lextureAllocationIndex[allocateddialogue],"class size:",dialogue[allocateddialogue][0])
                allocated = True
                return allocateddialogue
def main():
    studentarray = [] #Initialising local array database of student names and preferences
    studentcount = stucount(inputworksheet) #Retriving total no. of students
    
    for a in range(0,studentcount): #Initalising student array to accmodate for 3 values (name,preference1,preference2)
        studentarray.append([""]*3)
             
    for y in range(0,studentcount):#Copying values from input file to local database array
        studentarray[y][namecol-1]= inputworksheet.cell(row = y+1+header, column = namecol).value
        studentarray[y][choice1col-1]= inputworksheet.cell(row = y+1+header, column = choice1col).value
        studentarray[y][choice2col-1]= inputworksheet.cell(row = y+1+header, column = choice2col).value
    #if debug:
    #    for y in range(0,studentcount):
    #        print("Name:",inputworksheet.cell(row = y+1+header, column = 1).value,"Choice 1:",inputworksheet.cell(row = y+1+header, column = 2).value,"Choice 2:",inputworksheet.cell(row = y+1+header, column = 3).value)   
    #    print("\n")
    for b in range(0,studentcount): #Assigning classes to students + printing output
        if debug:
            print("Allocating student:",studentarray[b][0],"No.:",b)
        allocation = allocate(studentarray[b][1],studentarray[b][2])
        dialogueAllocation = allocateDialogue (studentarray[b][1],studentarray[b][2])
        wsoutput.cell(row=b+1+header,column=1).value = studentarray[b][0]
        wsoutput.cell(row=b+1+header,column=2).value = allocationIndex[allocation]
        wsoutput.cell(row=b+1+header,column=3).value = dialogueIndex[dialogueAllocation]
        if debug:
            print(studentarray[b][0],"is allocated to class",allocationIndex[allocation],"and dialogue: ",dialogueIndex[dialogueAllocation],"\n")
    
    wboutput.save("Result.xlsx")
    print("Done!")
    
    
main()
