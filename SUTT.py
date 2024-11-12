import pandas as pd
import json
try:
    data= pd.read_excel("Timetable Workbook - SUTT Task 1.xlsx",sheet_name=None,skiprows=1)
    print("File read successfully.")
except FileNotFoundError:
    print("File not found.")
file=[]
#traveling through each sheet for each object 

for i in range(1,7):
    sheet=data["S"+str(i)]

    #creating key value pairs for code and title
    fileobject={}
    fileobject["course_code"]=sheet.iloc[1,1]
    fileobject["course_title"]=sheet.iloc[1,2]
    
    #creating credits key value objects
    credits={}
    credits["lecture"]=sheet.iloc[1,3]
    credits["practical"]=sheet.iloc[1,4]
    credits["units"]=sheet.iloc[1,5]
    for j in credits:
        if credits[j]=="-":
            credits[j]=0
    fileobject["credits"]=credits

    #code for creating key value for section, room, timings
    sections=[]
    pos=1
    section=sheet.loc[pos,"SEC"]
    for k in range(1,sheet.count()["SEC"]+1):

        #code for creating sections key value pair
        sectionobject={}
        if section[0]=="P":
            sectiontype="practical"
        elif section[0]=="L":
            sectiontype="lecture"
        elif section[0]=="T":
            sectiontype="tutorial"
        sectionobject["section_type"]=sectiontype
        sectionobject["section_number"]=section
        sectionobject["room"]=str(int(sheet.loc[pos,"ROOM"]))

        #code for creating key value pair for timings
        timing=[]
        time=sheet.iloc[pos,9].split() 
        start=0
        for l in range(1,len(time)):
            if time[l].isdigit():
                end=l-1
                m=start
                while m<=end:
                    time[m]+=time[l]
                    m+=1
                if l+1<len(time) and time[l+1].isalpha():
                    start=l+1
        for n in time:
            if n[0].isalpha():
                timingobject={}
                timingobject["day"]=n[0]
                slots=[]
                if n[1].isalpha():
                    for o in n[2:]:
                        slots.append([int(o)+7,int(o)+8])
                else: 
                    for o in n[1:]:
                        slots.append([int(o)+7,int(o)+8])
                timingobject["slots"]=slots
                timing.append(timingobject)
        sectionobject["timing"]=timing
        
        #code for creating instructers key value pair
        instructors=[sheet.iloc[pos,7]]
        pos+=1
        if pos>=len(sheet):
                sectionobject["instructors"]=instructors
                sections.append(sectionobject)
                break
        section=str(sheet.loc[pos,"SEC"])
        while (section=="nan"):
            instructors.append(sheet.iloc[pos,7])
            pos+=1
            if pos>=len(sheet):
                break
            section=str(sheet.loc[pos,"SEC"])
        sectionobject["instructors"]=instructors
        
        sections.append(sectionobject)
    fileobject["sections"]=sections
    file.append(fileobject)
    print("Object for sheet",i,"created successfully.")
#dumping the created list into a json file
with open("sutt.json","w") as f:
    json.dump(file,f,indent=4)
    print("List of objects for each sheet added to the json file.")
    

