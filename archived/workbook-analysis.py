# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:percent
#     text_representation:
#       extension: .py
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.17.1
#   kernelspec:
#     display_name: Python 3 (ipykernel)
#     language: python
#     name: python3
# ---

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# FSM data processing notebook 
# done in the past, putting here to make more general and re-usable
# VERY OLD STUFF
##########################################################################################

import pandas as pd
import numpy as np
import sys
import os

# To run this notebook or any cells of this notebook you would need to make sure the data directory contains actual files

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process workbooks
##########################################################################################
cwd = os.getcwd()

rawWBChkData = os.path.join(cwd,'data/FSM/workbooks/Chuuk_BOY_2017-2018.xlsm')
rawWBKsaData = os.path.join(cwd,'data/FSM/workbooks/Kosrae_BOY_2017-2018.xlsm')
rawWBPniData = os.path.join(cwd,'data/FSM/workbooks/PNI_BOY_ 2017-2018.xlsm')
rawWBYapData = os.path.join(cwd,'data/FSM/workbooks/Yap_BOY_2017-2018.xlsm')

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
# %%time
# Load all four states worbooks in DataFrames
# (~20 seconds on iMac with i9 CPU and 32GB RAM)
chkSchoolsDF = pd.read_excel(rawWBChkData, sheet_name=['Schools'], header=2)['Schools']
chkStudentsDF = pd.read_excel(rawWBChkData, sheet_name=['Students'], header=2)['Students']
chkStaffsDF = pd.read_excel(rawWBChkData, sheet_name=['SchoolStaff'], header=2)['SchoolStaff']
chkWASHDF = pd.read_excel(rawWBChkData, sheet_name=['WASH'], header=2)['WASH']

ksaSchoolsDF = pd.read_excel(rawWBKsaData, sheet_name=['Schools'], header=2)['Schools']
ksaStudentsDF = pd.read_excel(rawWBKsaData, sheet_name=['Students'], header=2)['Students']
ksaStaffsDF = pd.read_excel(rawWBKsaData, sheet_name=['SchoolStaff'], header=2)['SchoolStaff']
ksaWASHDF = pd.read_excel(rawWBKsaData, sheet_name=['WASH'], header=2)['WASH']

pniSchoolsDF = pd.read_excel(rawWBPniData, sheet_name=['Schools'], header=2)['Schools']
pniStudentsDF = pd.read_excel(rawWBPniData, sheet_name=['Students'], header=2)['Students']
pniStaffsDF = pd.read_excel(rawWBPniData, sheet_name=['SchoolStaff'], header=2)['SchoolStaff']
pniWASHDF = pd.read_excel(rawWBPniData, sheet_name=['WASH'], header=2)['WASH']

yapSchoolsDF = pd.read_excel(rawWBYapData, sheet_name=['Schools'], header=2)['Schools']
yapStudentsDF = pd.read_excel(rawWBYapData, sheet_name=['Students'], header=2)['Students']
yapStaffsDF = pd.read_excel(rawWBYapData, sheet_name=['SchoolStaff'], header=2)['SchoolStaff']
yapWASHDF = pd.read_excel(rawWBYapData, sheet_name=['WASH'], header=2)['WASH']

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
print("chkSchools total: ", len(chkSchoolsDF))
print("chkStudents total: ", len(chkStudentsDF))
print("chkStaffs total: ", len(chkStaffsDF))
print("chkWASH total: ", len(chkWASHDF))
print("\n")

print("ksaSchools total: ", len(ksaSchoolsDF))
print("ksaStudents total: ", len(ksaStudentsDF))
print("ksaStaffs total: ", len(ksaStaffsDF))
print("ksaWASH total: ", len(ksaWASHDF))
print("\n")

print("pniSchools total: ", len(pniSchoolsDF))
print("pniStudents total: ", len(pniStudentsDF))
print("pniStaffs total: ", len(pniStaffsDF))
print("pniWASH total: ", len(pniWASHDF))
print("\n")

print("yapSchools total: ", len(yapSchoolsDF))
print("yapStudents total: ", len(yapStudentsDF))
print("yapStaffs total: ", len(yapStaffsDF))
print("yapWASH total: ", len(yapWASHDF))
print("\n")

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process schools
##########################################################################################
outSchoolsData = os.path.join(cwd,'data/FSM/workbooks/schools.xlsx')

#['SchoolYear','State','School Name', 'School No', 'First Name', 'Middle Name', 'Last Name', 'Full Name', 'Gender', 'Date of Birth', 'Age', 'Citizenship', 'Ethnicity', 'FSM_SSN', 
#'Highest Qualification', 'Field of Study', 'Year of Completion', 'Employment Status', 'Job Title', 'Organization', 'Staff Type', 'Date of Hire', 'Date Of Exit', 
#'ECE', 'Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 5', 'Grade 6', 'Grade 7', 'Grade 8', 'Grade 9', 'Grade 10', 'Grade 11', 'Grade 12','Admin'])

statesSchoolsDF = chkSchoolsDF.append([ksaSchoolsDF, pniSchoolsDF, yapSchoolsDF], ignore_index=True)
statesSchoolsDF

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process total schools by state
##########################################################################################
print("Schools Total: ", len(statesSchoolsDF))
print("Schools Chuuk Total: ", len(statesSchoolsDF[statesSchoolsDF['State'] == 'Chuuk']))
print("Schools Kosrae Total: ", len(statesSchoolsDF[statesSchoolsDF['State'] == 'Kosrae']))
print("Schools Pohnpei Total: ", len(statesSchoolsDF[statesSchoolsDF['State'] == 'Pohnpei']))
print("Schools Yap Total: ", len(statesSchoolsDF[statesSchoolsDF['State'] == 'Yap']))

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process teacher (all staff actually)
##########################################################################################
outTeachersData = os.path.join(cwd,'data/FSM/workbooks/teachers.xlsx')

#['SchoolYear','State','School Name', 'School No', 'First Name', 'Middle Name', 'Last Name', 'Full Name', 'Gender', 'Date of Birth', 'Age', 'Citizenship', 'Ethnicity', 'FSM_SSN', 
#'Highest Qualification', 'Field of Study', 'Year of Completion', 'Employment Status', 'Job Title', 'Organization', 'Staff Type', 'Date of Hire', 'Date Of Exit', 
#'ECE', 'Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 5', 'Grade 6', 'Grade 7', 'Grade 8', 'Grade 9', 'Grade 10', 'Grade 11', 'Grade 12','Admin'])

statesTeachersDF = chkStaffsDF.append([ksaStaffsDF, pniStaffsDF, yapStaffsDF], ignore_index=True)

statesTeachersDF

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process students
##########################################################################################
outStudentsData = os.path.join(cwd,'data/FSM/workbooks/students.xlsx')

statesStudentsDF = chkStudentsDF.append([ksaStudentsDF, pniStudentsDF, yapStudentsDF], ignore_index=True)

statesStudentsDF

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process students in elementary and secondary of official age by state (net enrolments)
##########################################################################################

statesStudentsDF1 = statesStudentsDF[
[
 'First Name',
 'Last Name',
 'State',
 'Grade Level',
 'Age']
]

print("States Students for SY2017-18: ", statesStudentsDF1.head(3))
print("Total Students for SY2017-18: ", len(statesStudentsDF))
print("Grade Level: ", statesStudentsDF1['Grade Level'].unique())
ages = statesStudentsDF1['Age'].unique()
ages.sort()
print("Ages: ", ages)
print("State: ", statesStudentsDF1['State'].unique())
print("\n")

statesNetPrimaryStudentsDF = statesStudentsDF1[
((statesStudentsDF1['Grade Level'] == 'Grade ECE') & (statesStudentsDF1['Age'] == 5)) |
((statesStudentsDF1['Grade Level'] == 'Grade 1') & (statesStudentsDF1['Age'] == 6)) |
((statesStudentsDF1['Grade Level'] == 'Grade 2') & (statesStudentsDF1['Age'] == 7)) |
((statesStudentsDF1['Grade Level'] == 'Grade 3') & (statesStudentsDF1['Age'] == 8)) |
((statesStudentsDF1['Grade Level'] == 'Grade 3 ') & (statesStudentsDF1['Age'] == 8)) |
((statesStudentsDF1['Grade Level'] == 'Grade 4') & (statesStudentsDF1['Age'] == 9)) |
((statesStudentsDF1['Grade Level'] == 'Grade 5') & (statesStudentsDF1['Age'] == 10)) |
((statesStudentsDF1['Grade Level'] == 'Grade 6') & (statesStudentsDF1['Age'] == 11)) |
((statesStudentsDF1['Grade Level'] == 'Grade 7') & (statesStudentsDF1['Age'] == 12)) |
((statesStudentsDF1['Grade Level'] == 'Grade 7 ') & (statesStudentsDF1['Age'] == 12)) |
((statesStudentsDF1['Grade Level'] == 'Grade 8') & (statesStudentsDF1['Age'] == 13))]

#print("Primary Students: ", statesNetPrimaryStudentsDF.count())
print("Net Primary Students Total: ", len(statesNetPrimaryStudentsDF))
#print("Primary Students Chuuk: ", statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Chuuk'].count())
print("Net Primary Students Chuuk Total: ", len(statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Chuuk']))
#print("Primary Students Kosrae: ", statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Kosrae'].count())
print("Net Primary Students Kosrae Total: ", len(statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Kosrae']))
#print("Primary Students Pohnpei: ", statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Pohnpei'].count())
print("Net Primary Students Pohnpei Total: ", len(statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Pohnpei']))
#print("Primary Students Yap: ", statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Yap'].count())
print("Net Primary Students Yap Total: ", len(statesNetPrimaryStudentsDF[statesNetPrimaryStudentsDF['State'] == 'Yap']))

statesNetSecondaryStudentsDF = statesStudentsDF1[
((statesStudentsDF1['Grade Level'] == 'Grade 9') & (statesStudentsDF1['Age'] == 14)) |
((statesStudentsDF1['Grade Level'] == 'Grade 10') & (statesStudentsDF1['Age'] == 15)) |
((statesStudentsDF1['Grade Level'] == 'Grade 11') & (statesStudentsDF1['Age'] == 16)) |
((statesStudentsDF1['Grade Level'] == 'Grade 12') & (statesStudentsDF1['Age'] == 17))]

#print("Secondary Students: ", statesNetSecondaryStudentsDF.count())
print("Net Secondary Students Total: ", len(statesNetSecondaryStudentsDF))
#print("Secondary Students Chuuk: ", statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Chuuk'].count())
print("Net Secondary Students Chuuk Total: ", len(statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Chuuk']))
#print("Secondary Students Kosrae: ", statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Kosrae'].count())
print("Net Secondary Students Kosrae Total: ", len(statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Kosrae']))
#print("Secondary Students Pohnpei: ", statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Pohnpei'].count())
print("Net Secondary Students Pohnpei Total: ", len(statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Pohnpei']))
#print("Secondary Students Yap: ", statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Yap'].count())
print("Net Secondary Students Yap Total: ", len(statesNetSecondaryStudentsDF[statesNetSecondaryStudentsDF['State'] == 'Yap']))



# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process active students in elementary and secondary by state (total enrolments)
##########################################################################################

statesStudentsDF1 = statesStudentsDF[
[
 'First Name',
 'Last Name',
 'State',
 'Grade Level',
 'Age']
]

print("States Students: ", statesStudentsDF1.head(3))
print("Grade Level: ", statesStudentsDF1['Grade Level'].unique())
print("State: ", statesStudentsDF1['State'].unique())
print("\n")

statesActivePrimaryStudentsDF = statesStudentsDF1[
(statesStudentsDF1['Grade Level'] == 'Grade ECE') |
(statesStudentsDF1['Grade Level'] == 'Grade 1') |
(statesStudentsDF1['Grade Level'] == 'Grade 2') |
(statesStudentsDF1['Grade Level'] == 'Grade 3') |
(statesStudentsDF1['Grade Level'] == 'Grade 3 ') |
(statesStudentsDF1['Grade Level'] == 'Grade 4') |
(statesStudentsDF1['Grade Level'] == 'Grade 5') |
(statesStudentsDF1['Grade Level'] == 'Grade 6') |
(statesStudentsDF1['Grade Level'] == 'Grade 7') |
(statesStudentsDF1['Grade Level'] == 'Grade 7 ') |
(statesStudentsDF1['Grade Level'] == 'Grade 8')]

#print("Primary Students: ", statesActivePrimaryStudentsDF.count())
print("Primary Students Total: ", len(statesActivePrimaryStudentsDF))
#print("Primary Students Chuuk: ", statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Chuuk'].count())
print("Primary Students Chuuk Total: ", len(statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Chuuk']))
#print("Primary Students Kosrae: ", statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Kosrae'].count())
print("Primary Students Kosrae Total: ", len(statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Kosrae']))
#print("Primary Students Pohnpei: ", statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Pohnpei'].count())
print("Primary Students Pohnpei Total: ", len(statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Pohnpei']))
#print("Primary Students Yap: ", statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Yap'].count())
print("Primary Students Yap Total: ", len(statesActivePrimaryStudentsDF[statesActivePrimaryStudentsDF['State'] == 'Yap']))

statesActiveSecondaryStudentsDF = statesStudentsDF1[
(statesStudentsDF1['Grade Level'] == 'Grade 9') |
(statesStudentsDF1['Grade Level'] == 'Grade 10') |
(statesStudentsDF1['Grade Level'] == 'Grade 11') |
(statesStudentsDF1['Grade Level'] == 'Grade 12')]

#print("Secondary Students: ", statesActiveSecondaryStudentsDF.count())
print("Secondary Students Total: ", len(statesActiveSecondaryStudentsDF))
#print("Secondary Students Chuuk: ", statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Chuuk'].count())
print("Secondary Students Chuuk Total: ", len(statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Chuuk']))
#print("Secondary Students Kosrae: ", statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Kosrae'].count())
print("Secondary Students Kosrae Total: ", len(statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Kosrae']))
#print("Secondary Students Pohnpei: ", statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Pohnpei'].count())
print("Secondary Students Pohnpei Total: ", len(statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Pohnpei']))
#print("Secondary Students Yap: ", statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Yap'].count())
print("Secondary Students Yap Total: ", len(statesActiveSecondaryStudentsDF[statesActiveSecondaryStudentsDF['State'] == 'Yap']))


# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process teachers teaching Grade 5 and 7
##########################################################################################

statesTeachersDF2 = statesTeachersDF.copy()

#statesTeachersDF = statesTeachersDF[(statesTeachersDF['Staff Type'] == 'Teaching Staff')]
statesTeachersDF2 = statesTeachersDF2[
    (statesTeachersDF2['Staff Type'] == 'Teaching Staff') & 
    ((statesTeachersDF2['Grade 5'] == 'x') | 
     (statesTeachersDF2['Grade 7'] == 'x') | 
     (statesTeachersDF2['Grade 5'] == 'X') | 
     (statesTeachersDF2['Grade 7'] == 'X'))]

statesTeachersDF2 = statesTeachersDF2.drop(
    ['School Type', 'Office', 'OTHER_SSN', 'Highest Ed Qualification', 'Year Of Completion2', 
     'Reason', 'Teacher-Type', 'Annual Salary', 'Funding Source', 'Other', 'Total Days Absence', 
     'Admin', 'Maths', 'Science', 'Language', 'Competency', 'Age', 'Citizenship', 'Date Of Exit', 
     'Date of Birth', 'Ethnicity', 'Field of STudy', 'Field of Study', 'Organization', 
     'Year of Completion', 'ECE', 'Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 6', 
     'Grade 8', 'Grade 9', 'Grade 10', 'Grade 11', 'Grade 12'], inplace=False, axis='columns') # , 'Staff Type'

statesTeachersDF2 = statesTeachersDF2.reindex(
    columns = ['State','School Name', 'School No', 'SchoolYear', 'Full Name', 'First Name', 
               'Middle Name', 'Last Name', 'Gender', 'FSM_SSN', 'Date of Hire', 
               'Employment Status', 'Grade 5', 'Grade 7', 'Highest Qualification', 
               'Job Title', 'Staff Type'])


# Writing data to sheets
writer = pd.ExcelWriter(outTeachersData)
statesTeachersDF2.to_excel(writer, sheet_name='Teachers', index=False)
writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process active teachers in elementary and secondary by state
##########################################################################################

statesTeachersDF3 = statesTeachersDF.copy()

statesTeachersDF3 = statesTeachersDF3[
[
 #'Date of Birth',
 #'SchoolYear', 
 'First Name',
 'Last Name',
 'Employment Status',
 'Staff Type',
 'State',
 'Teacher-Type',
 'ECE',
 #'Gender',
 'Grade 1',
 'Grade 10',
 'Grade 11',
 'Grade 12',
 'Grade 2',
 'Grade 3',
 'Grade 4',
 'Grade 5',
 'Grade 6',
 'Grade 7',
 'Grade 8',
 'Grade 9',
 #'Highest Ed Qualification',
 #'Highest Qualification',
 #'Job Title',
 #'Middle Name',
 #'School Name',
 #'School No',
 ]
]

print("Employment Status: ", statesTeachersDF3['Employment Status'].unique())
print("Staff Type: ", statesTeachersDF3['Staff Type'].unique())
print("State: ", statesTeachersDF3['State'].unique())
print("Teacher-Type: ", statesTeachersDF3['Teacher-Type'].unique())
print("Grades: ", pd.unique(statesTeachersDF3[['ECE', 'Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 5', 'Grade 6', 'Grade 7', 'Grade 8', 'Grade 9', 'Grade 10', 'Grade 11', 'Grade 12']].values.ravel('K')))
print("\n")

statesActivePrimaryTeachersDF = statesTeachersDF3[
(statesTeachersDF3['Employment Status'] == 'Active') &
(statesTeachersDF3['Staff Type'] == 'Teaching Staff') &
(((statesTeachersDF3['ECE'] == 'X') | (statesTeachersDF3['ECE'] == 'x')) |
((statesTeachersDF3['Grade 1'] == 'X') | (statesTeachersDF3['Grade 1'] == 'x')) |
((statesTeachersDF3['Grade 2'] == 'X') | (statesTeachersDF3['Grade 2'] == 'x')) |
((statesTeachersDF3['Grade 3'] == 'X') | (statesTeachersDF3['Grade 3'] == 'x')) |
((statesTeachersDF3['Grade 4'] == 'X') | (statesTeachersDF3['Grade 4'] == 'x')) |
((statesTeachersDF3['Grade 5'] == 'X') | (statesTeachersDF3['Grade 5'] == 'x')) |
((statesTeachersDF3['Grade 6'] == 'X') | (statesTeachersDF3['Grade 6'] == 'x')) |
((statesTeachersDF3['Grade 7'] == 'X') | (statesTeachersDF3['Grade 7'] == 'x')) |
((statesTeachersDF3['Grade 8'] == 'X') | (statesTeachersDF3['Grade 8'] == 'x')))]

#print("Primary Teachers: ", statesActivePrimaryTeachersDF.count())
print("Primary Teachers Total: ", len(statesActivePrimaryTeachersDF))
#print("Primary Teachers Chuuk: ", statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Chuuk'].count())
print("Primary Teachers Chuuk Total: ", len(statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Chuuk']))
#print("Primary Teachers Kosrae: ", statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Kosrae'].count())
print("Primary Teachers Kosrae Total: ", len(statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Kosrae']))
#print("Primary Teachers Pohnpei: ", statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Pohnpei'].count())
print("Primary Teachers Pohnpei Total: ", len(statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Pohnpei']))
#print("Primary Teachers Yap: ", statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Yap'].count())
print("Primary Teachers Yap Total: ", len(statesActivePrimaryTeachersDF[statesActivePrimaryTeachersDF['State'] == 'Yap']))

statesActiveSecondaryTeachersDF = statesTeachersDF3[
(statesTeachersDF3['Employment Status'] == 'Active') &
(statesTeachersDF3['Staff Type'] == 'Teaching Staff') &
(((statesTeachersDF3['Grade 9'] == 'X') | (statesTeachersDF3['Grade 9'] == 'x')) |
((statesTeachersDF3['Grade 10'] == 'X') | (statesTeachersDF3['Grade 10'] == 'x')) |
((statesTeachersDF3['Grade 11'] == 'X') | (statesTeachersDF3['Grade 11'] == 'x')) |
((statesTeachersDF3['Grade 12'] == 'X') | (statesTeachersDF3['Grade 12'] == 'x')))]

#print("Secondary Teachers: ", statesActiveSecondaryTeachersDF.count())
print("Secondary Teachers Total: ", len(statesActiveSecondaryTeachersDF))
#print("Secondary Teachers Chuuk: ", statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Chuuk'].count())
print("Secondary Teachers Chuuk Total: ", len(statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Chuuk']))
#print("Secondary Teachers Kosrae: ", statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Kosrae'].count())
print("Secondary Teachers Kosrae Total: ", len(statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Kosrae']))
#print("Secondary Teachers Pohnpei: ", statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Pohnpei'].count())
print("Secondary Teachers Pohnpei Total: ", len(statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Pohnpei']))
#print("Secondary Teachers Yap: ", statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Yap'].count())
print("Secondary Teachers Yap Total: ", len(statesActiveSecondaryTeachersDF[statesActiveSecondaryTeachersDF['State'] == 'Yap']))

# Writing data to sheets
# writer = pd.ExcelWriter(outTeachersData)
# statesTeachersDF.to_excel(writer, sheet_name='Teachers', index=False)
# writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
print("Students Total: ", len(statesStudentsDF1))
print("Students Chuuk Total: ", len(statesStudentsDF1[statesStudentsDF1['State'] == 'Chuuk']))
print("Students Kosrae Total: ", len(statesStudentsDF1[statesStudentsDF1['State'] == 'Kosrae']))
print("Students Pohnpei Total: ", len(statesStudentsDF1[statesStudentsDF1['State'] == 'Pohnpei']))
print("Students Yap Total: ", len(statesStudentsDF1[statesStudentsDF1['State'] == 'Yap']))


# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
# Extra data directly from FedEMIS

# SY2016-17 (this quite closely lines up with previously given data in CombinedData spreadsheet)
# Elementary => 20161
# Secondary => 6214
# Chuuk Elementary => 904 + 8350 = 9254
# Chuuk Secondary => 2328
# Kosrae Elementary => 188 + 1236 = 1424
# Kosrae Secondary => 642
# Pohnpei Elementary => 659 + 6586 = 7245
# Pohnpei Secondary => 2383
# Yap Elementary => 388 + 1850 = 2238
# Yap Secondary => 861

# SY2017-18 (this does not line up with data from workbook, some duplicates in there? missing data?)
# Elementary => 20228
# Secondary => 6354

# From workbook manually to validate my numbers in previous cells (this matches and therefore my pandas processing seems correct)
# State CountFromWorkbook (Summed from processed figures above)
# Chuuk 10839 (10839)
# Kosrae 2013 (2013)
# Yap 2911 (2911)
# Pohnei 10307 (10307)
# Total 26070 (26070)

# First my pandas processed number then the number from XML vermpaf
#Primary Students Total:  19737
#Primary Students Chuuk Total:  8613 (10279)
#Primary Students Kosrae Total:  1382 (1382)
#Primary Students Pohnpei Total:  7660 (7594)
#Primary Students Yap Total:  2082 (2156)
#Secondary Students Total:  6333
#Secondary Students Chuuk Total:  2226 (2237)
#Secondary Students Kosrae Total:  631 (630)
#Secondary Students Pohnpei Total:  2647 (2657)
#Secondary Students Yap Total:  829 (830)

# %%
