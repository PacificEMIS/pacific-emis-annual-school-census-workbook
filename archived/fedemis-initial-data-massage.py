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
# FSM data cleanup, processing notebook
# VERY OLD STUFF. Temporary home if ever need to review, clean and make re-usable and better
# It was used to initial cleanup, massage and prepare data for loading in FedEMIS
##########################################################################################

import pandas as pd
import numpy as np
import sys
import os

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process population projections (2010 Census)
##########################################################################################
cwd = os.getcwd()

rawPopProjData = os.path.join(cwd,'data/FSM/PopulationProjection2010CensusProcessed.xlsx')
outPopProjData = os.path.join(cwd,'data/FSM/fsm-population-10-year-projection-initial-data.xlsx')

usecols = {
    2010: "A:C",
    2011: "F:H",
    2012: "K:M",
    2013: "P:R",
    2014: "U:W",
    2015: "Z:AB",
    2016: "AE:AG",
    2017: "AJ:AL",
    2018: "AO:AQ"
}

sheets = {
    "Chuuk": [1,"CHK"],
    "Pohnpei": [3, "PNI"],
    "Yap": [4, "YAP"],
    "Kosrae": [2, "KSA"]
}

data = {}

for key, value in usecols.items():
    data[key] = pd.read_excel(rawPopProjData, sheet_name=None, header=2, index_col=None,
                              names=["popAge","popM","popF"], usecols=value)

popProjDF = pd.DataFrame()

# for each year
for year in data:
    # for each state
    for state in data[year]:
        # I'm here got a DataFrame
        #print(y,s)
        df = data[year][state]
        df['popmodCode'] = pd.Series(["FSMNSO"] * 76, index=df.index)
        df['dID'] = pd.Series([sheets[state][0]] * 76, index=df.index)
        df['elN'] = pd.Series([sheets[state][1]] * 76, index=df.index)
        df['popYear'] = pd.Series([year] * 76, index=df.index)
        popProjDF = popProjDF.append(df)

popProjDF = popProjDF[["popmodCode","popYear","popAge","popM","popF","dID","elN"]]
popProjDF.loc[popProjDF.popAge == '        75+', 'popAge'] = 75
popProjDF = popProjDF.round({'popM': 0, 'popF': 0})
popProjDF

# Experiment with population projection data for quick and dirty
# quality test
#data[2010]["Pohnpei"]
#df2013 = popProjDF[popProjDF['popYear'] == 2013]
#df2013.sum()

# Write population data
popProjDF.to_excel(outPopProjData, index=False)

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Process population projections (2010 Census)
##########################################################################################
import numpy as np

rawPopProjData3 = os.path.join(cwd,'data/FSM/Pop Projection (HIES 13-based)1.xlsx')
outPopProjData3 = os.path.join(cwd,'data/FSM/fsm-population-10-year-projection-initial-data3.xlsx')

popProjDataDF3 = pd.read_excel(rawPopProjData3, sheet_name=None, header=[1,2], index_col=0)

popProjDataDF3 = popProjDataDF3['HIES13 Pop Proj']

# cleanup indices
popProjDataDF3.rename(columns={
    '2010A' : 2010,
    '2013B': 2013,
    '2014': 2014,
    '2015': 2015,
    '2016': 2016,
    '2017': 2017,
    '2018': 2018,
    '2019': 2019,
    '2020': 2020,
    '2021': 2021,
    '2022': 2022,
    '2023': 2023,
    '2024': 2024,
    '2025': 2025
}, inplace=True)
popProjDataDF3.rename_axis(['Year', 'Gender'], axis="columns", inplace=True)

# group the DataFrame
popProjDataDF3Total = popProjDataDF3[0:80]
popProjDataDF3Yap = popProjDataDF3[80:160]
popProjDataDF3Chuuk = popProjDataDF3[160:240]
popProjDataDF3Pohnpei = popProjDataDF3[240:320]
popProjDataDF3Kosrae = popProjDataDF3[320:400]

# Add a region index
popProjDataDF3Total = popProjDataDF3Total.assign(Region=pd.Series(np.nan))
popProjDataDF3Total.set_index('Region', append=True, inplace=True)
popProjDataDF3Yap = popProjDataDF3Yap.assign(Region='YAP')
popProjDataDF3Yap.set_index('Region', append=True, inplace=True)
popProjDataDF3Chuuk = popProjDataDF3Chuuk.assign(Region='CHK')
popProjDataDF3Chuuk.set_index('Region', append=True, inplace=True)
popProjDataDF3Pohnpei = popProjDataDF3Pohnpei.assign(Region='PNI')
popProjDataDF3Pohnpei.set_index('Region', append=True, inplace=True)
popProjDataDF3Kosrae = popProjDataDF3Kosrae.assign(Region='KSA')
popProjDataDF3Kosrae.set_index('Region', append=True, inplace=True)

# Put the DataFrames back together (don't really need the popProjDataDF3Total do we)
popProjDataDF3 = pd.concat([popProjDataDF3Yap, popProjDataDF3Chuuk, popProjDataDF3Pohnpei, popProjDataDF3Kosrae])

# Bit more indices cleanup and remove unnecessary rows
popProjDataDF3.rename_axis(['Age', 'Region'], axis='index', inplace=True)
popProjDataDF3 = popProjDataDF3.reorder_levels(['Region', 'Age'])
#print(popProjDataDF3.index)

popProjDataDF3 = popProjDataDF3.drop(index=[
    'Yap','Chuuk','Pohnpei','Kosrae',
    '0 to 4', '5 to 9', '10 to 14', '15 to 19', '20 to 24', '25 to 29', '30 to 34', '35 to 39', '40 to 44',
    '45 to 49', '50 to 54', '55 to 59', '60 to 64', '65+'], level=1)

popProjDataDF3 = popProjDataDF3.reset_index()
popProjDataDF3['Age'] = popProjDataDF3['Age'].astype('int64')
popProjDataDF3 = popProjDataDF3.set_index(['Region', 'Age'])
popProjDataDF3.sort_index(axis=1, level=0, inplace=True)
#print(popProjDataDF3.index)

# Leave this rounding to excel
#popProjDataDF3 = popProjDataDF3.apply(np.int64)
#popProjDataDF3.info()

# process and flatten index
columns_mi = popProjDataDF3.columns
ind = pd.Index([str(e[1]) + str(e[0]) for e in columns_mi.tolist()])
ind
popProjDataDF3.columns = ind
# print(popProjDataDF3)
# print("\n")

# Set index as a column for use with wide_to_long
popProjDataDF3["id"] = popProjDataDF3.index
# print(popProjDataDF3.head(2))
# print("\n")

popProjDataDF3 = pd.wide_to_long(popProjDataDF3, ["Female", "Male", "Total"], i="id", j="popYear")
# cleanup indexes (move year to column, change back tuple Index into MultiIndex)
popProjDataDF3.reset_index(level='popYear', inplace=True)
popProjDataDF3.index = pd.MultiIndex.from_tuples(popProjDataDF3.index, names=['Region', 'Age'])
popProjDataDF3.reset_index(inplace=True)

# Final rename and missing columns
popProjDataDF3.rename(columns={'Female': 'popF', 'Male': 'popM', 'Region': 'dID', 'Age': 'popAge'}, inplace=True)
popProjDataDF3['elN'] = popProjDataDF3['dID']
popProjDataDF3 = popProjDataDF3.drop(['Total'], 1)
popProjDataDF3['popmodCode']  = 'FSMNSO2'
order_cols = ['popmodCode','popYear','popAge','popM','popF','dID','elN']
popProjDataDF3 = popProjDataDF3[order_cols]

print(popProjDataDF3)

# Write population data
popProjDataDF3.to_excel(outPopProjData3, index=False)

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Used in previous code, no longer needed as it was done already and saved in spreadsheet,
# skip to next cell for short
##########################################################################################

# rawSchoolsExtraData = os.path.join(cwd,'data/FSM/All_FSM_Schools.xlsx')
# rawInitialSchoolsData= os.path.join(cwd,'data/FSM/fsm-schools-with-raw-from-andrew-and-weison.xlsx')
# outInitialSchoolsTyposManualData = os.path.join(cwd,'data/FSM/fsm-schools-initial-lookups-anomalies-manual.xlsx')
# outInitialSchoolsTyposData = os.path.join(cwd,'data/FSM/fsm-schools-initial-lookups-anomalies.xlsx')
# outSchoolsLookup = os.path.join(cwd,'data/FSM/fsm-schools-lookups.xlsx')


# # Prepare schools helpers lookups. This is meant to assign the correct schools from badly entered data in various spreadsheets, etc.
# #schoolsLookupDF = pd.read_excel(outEnrolmentTransitData, sheet_name="cleanMergedSchoolsLeftJoinDF", header=0, usecols="A,B")
# schoolsLookupDF = pd.read_excel(rawInitialSchoolsData, sheet_name="Schools", header=0, usecols="A,B")
# schoolsLookupTyposFromEnrolmentsDF = pd.read_excel(outInitialSchoolsTyposManualData, sheet_name="FromEnrolments", header=0, usecols="A,B")
# schoolsLookupTyposFromTeachersDF = pd.read_excel(outInitialSchoolsTyposManualData, sheet_name="FromTeachers", header=0, usecols="A,B")
# schoolsLookupTyposFromTeachersDF = schoolsLookupTyposFromTeachersDF.rename(columns = {'RawSchools': 'schName','SchNoMappingCompleteMe': 'schNo'})
# schoolsLookupTyposFromAccreditationsDF = pd.read_excel(outInitialSchoolsTyposManualData, sheet_name="FromAccreditations", header=0, usecols="A,B")
# schoolsLookupTyposFromAccreditationsDF = schoolsLookupTyposFromAccreditationsDF.rename(columns = {'RawSchools': 'schName','SchNoMappingCompleteMe': 'schNo'})
# # remove duplicates
# schoolsLookupTyposFromTeachersDF.drop_duplicates(['schName','schNo'], inplace=True)
# schoolsLookupTyposFromAccreditationsDF.drop_duplicates(['schName','schNo'], inplace=True)

# # Add the typos constructed lookups containing to new manually fixed schools lookup
# schoolsLookupDF = schoolsLookupDF.append(schoolsLookupTyposFromEnrolmentsDF)
# schoolsLookupDF = schoolsLookupDF.append(schoolsLookupTyposFromTeachersDF)
# schoolsLookupDF = schoolsLookupDF.append(schoolsLookupTyposFromAccreditationsDF)
# schoolsLookupDF.drop_duplicates(['schName'], inplace=True)
# # remove the ones without an key (NaN means they have a typo of some sort)
# schoolsLookupDF = schoolsLookupDF.dropna()
# print('schoolsLookupDF:')
# print(schoolsLookupDF)

# # Final dataset containing a correct school ID for each school names in various
# # spreadsheets including all the ones with typos and differently spelled
# schoolsLookup = schoolsLookupDF.set_index('schName').to_dict()['schNo']
# schoolsLookupByName = {y:x for x,y in schoolsLookup.items()}
# print('schoolsLookup:')
# print(schoolsLookup)

# # Writing data to sheets
# writer = pd.ExcelWriter(outSchoolsLookup)
# schoolsLookupDF.to_excel(writer, sheet_name='SchoolsLookups', index=False)
# writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Create a comprehensive school lookup map with both old and new school ID
##########################################################################################

# Starting point of the old lookup containing inconsistencies from
# teacher, accreditation and enrollment raw data
oldSchoolLookupData = os.path.join(cwd,'data/FSM/fsm-schools-lookups-old.xlsx')
oldSchoolLookupDataDF = pd.read_excel(oldSchoolLookupData, header=0)
oldSchoolLookupDataDF.insert(1,'newSchNo',None)
print(oldSchoolLookupDataDF[:3])

# Now new mappings from Weison/Eugene
newSchoolIDLookupRawData = os.path.join(cwd,'data/FSM/2017 UpdateSchool_Code with old Codes.xlsx')

newSchoolIDLookupDict = {}
usecols = {
    'chuuk': "A:C",
    'yap': "F:H",
    'pohnpei': "K:M",
    'kosrae': "P:R",
}

for key, value in usecols.items():
    newSchoolIDLookupDict[key] = pd.read_excel(newSchoolIDLookupRawData, header=1, index_col=None,
                              names=["schName","newSchNo","schNo"], usecols=value)

newSchoolIDLookupDataDF = pd.concat(newSchoolIDLookupDict)
newSchoolIDLookupDataDF = newSchoolIDLookupDataDF.dropna(axis='index', how='any')
newSchoolIDLookupDataDF = newSchoolIDLookupDataDF.reset_index(drop=True)
newSchoolIDLookupDataDF

# Fill old mapping with new school ID
oldToNewDict = pd.Series(newSchoolIDLookupDataDF.newSchNo.values,index=newSchoolIDLookupDataDF.schNo).to_dict()
oldSchoolLookupDataDF.newSchNo = oldSchoolLookupDataDF.newSchNo.fillna(oldSchoolLookupDataDF.schNo.map(oldToNewDict))
print(oldSchoolLookupDataDF[:3])


# Combined all mapping together
schoolLookupMappingDF = oldSchoolLookupDataDF.append(newSchoolIDLookupDataDF,ignore_index=True)
schoolLookupMappingDF

schoolsLookup = schoolLookupMappingDF.set_index('schName').to_dict()['newSchNo']
schoolsLookupByName = {y:x for x,y in schoolsLookup.items()}
print('schoolsLookup:')
print(schoolsLookup)

# Write school lookup mapping for later use (i.e. processing enrollments, teachers, accredication raw data)
newSchoolLookupData = os.path.join(cwd,'data/FSM/fsm-schools-lookups.xlsx')
schoolLookupMappingDF.to_excel(newSchoolLookupData, index=False)

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
##########################################################################################
# Get some offical lookups for use in various data processing throughout cells
##########################################################################################

rawFSMLookups = os.path.join(cwd,'data/FSM/PineapplesLookups-FSM.xlsx')

authoritiesLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="Authorities", skiprows=38, usecols="A,B")
authoritiesLookups = authoritiesLookupsDF.set_index('authName').to_dict()['authCode']
authoritiesByNameLookups = authoritiesLookupsDF.set_index('authCode').to_dict()['authName']
print('authoritiesLookupsDF:')
print(authoritiesLookupsDF.head(15))

gradeLevelsLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="lkpLevels", skiprows=44, usecols="A,B")
print('gradeLevelsLookupsDF:')
print(gradeLevelsLookupsDF.head(5))

electLLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="lkpElectorateL", skiprows=85, usecols="A,B")
print('electLLookupsDF:')
print(electLLookupsDF.head(5))

electNLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="lkpElectorateN", skiprows=36, usecols="A,B")
print('electNLookupsDF:')
print(electNLookupsDF.head(5))

# islands and municipalities
islandsLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="Islands", skiprows=94, usecols="A,B")
print('islandsLookupsDF:')
print(islandsLookupsDF.head(5))
islandsLookup = islandsLookupsDF.set_index('iName').to_dict()['iCode']
islandsLookupByName = {y:x for x,y in islandsLookup.items()}
print('islandsLookup:')
print(islandsLookup)


teacherRoleLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="lkpTeacherRole", skiprows=116, usecols="A,B")
teacherRoleLookups = teacherRoleLookupsDF.set_index('codeDescription').to_dict()['codeCode']
teacherRoleByNameLookups = teacherRoleLookupsDF.set_index('codeCode').to_dict()['codeDescription']
print('teacherRoleLookupsDF:')
print(teacherRoleLookupsDF.head(5))

teacherQualLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="lkpTeacherQual", skiprows=55, usecols="A,B")
print('teacherQualLookupsDF:')
print(teacherQualLookupsDF.head(5))

roleGradesLookupsDF = pd.read_excel(rawFSMLookups, sheet_name="RoleGrades", skiprows=82, usecols="A,C")
roleGradesLookups = roleGradesLookupsDF.set_index('roleCode').to_dict()['rgCode']
print('roleGradesLookupsDF:')
print(roleGradesLookupsDF.head(5))

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
################################################################################
# Process schools data
################################################################################

rawSchoolsData = os.path.join(cwd,'data/FSM/schools-cleaned-from-raw-data.xlsx')
schoolsOrderedDictDF = pd.read_excel(rawSchoolsData, sheet_name=['Chuuk','Yap','Pohnpei','Kosrae'], header=0)
schoolsDF = pd.concat([schoolsOrderedDictDF['Chuuk'],schoolsOrderedDictDF['Yap'],schoolsOrderedDictDF['Pohnpei'],schoolsOrderedDictDF['Kosrae']])
schoolsDF = schoolsDF.reset_index(drop=True)

# Drop columns with lesser quality data, bad data or uneeded data
schoolsDF = schoolsDF.drop(['schAuth','schElectL','Location','Region/Zone/Municipality','School Type','School Level','Enrollment'], 1)

# Map authorities to their codes
schoolsDF = schoolsDF.replace(to_replace={'Authority':authoritiesLookups})
schoolsDF = schoolsDF.rename(columns = {'Authority': 'schAuth'})

# Map Islands/Municipality (i.e. iCode) to their codes
schoolsDF = schoolsDF.replace(to_replace={'iCode':islandsLookup})

# Use 'IslandsOrElectorate' to infer Local electorate (i.e. schElectL)
def assignElectLOrNull(x):
    # will get iCode another way
    # if x in islandsLookupsDF['iName'].values:        
    #     # return the iCode
    #     return islandsLookupsDF[islandsLookupsDF['iName'] == x]['iCode'].values[0]
    if x in electLLookupsDF['codeDescription'].values:
        # return the codeCode
        return electLLookupsDF[electLLookupsDF['codeDescription'] == x]['codeCode'].values[0]
    else:
        return 'NULL'

schoolsDF = schoolsDF.assign(schElectL = schoolsDF['IslandsOrElectorate'])
schoolsDF['schElectL'] = schoolsDF['schElectL'].apply(assignElectLOrNull)
schoolsDF = schoolsDF.drop(['IslandsOrElectorate'], 1)

# Put NULL for all missing values
schoolsDF = schoolsDF.fillna('NULL')

# Set registration date for all schools to 1 Jan 1970 for now
from datetime import date
schoolsDF = schoolsDF.assign(schRegStatusDate=date(1970,1,1))

# Replace all NULL with 0 in schClosed and replace 1 with 2016 (assumed year school closed)
schoolsDF.loc[schoolsDF.schClosed == 'NULL', 'schClosed'] = 0
schoolsDF.loc[schoolsDF.schClosed == 1, 'schClosed'] = 2016

# Handle any school name with ' in it.
schoolsDF['schName'] = schoolsDF['schName'].str.replace("'","''")

print('schoolsDF:')
print(schoolsDF.head(5))

# Writing data to sheets
outSchoolsInitialData = os.path.join(cwd,'data/FSM/fsm-schools-initial-data.xlsx')
writer = pd.ExcelWriter(outSchoolsInitialData)
schoolsDF.to_excel(writer, sheet_name='Schools', index=False)
writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "slide"}
# %%time
################################################################################
# Prepare student enrolments auxiliary raw data
# Also get all schools to identify typos and associate with correct school (in other cell)
# And some grade level data processing too
# (~50 seconds on iMac with i9 CPU and 32GB RAM)
################################################################################

rawCombinedData = os.path.join(cwd,'data/FSM/CombinedData.xlsx')

# Read the whole Student Roster in
rawRosterDF = pd.read_excel(rawCombinedData, sheet_name="CombinedStudents", header=0, index_col=None, usecols="A,C,I,J,M,N,P,S,T,U,X")
print("rawRosterDF: ")
print(rawRosterDF.head(1))

# Prepare grade levels helper lookup (from student roster)
uniqueLevelsInStudentRoster = rawRosterDF['Grade Level'].unique()
uniqueLevelsInStudentRosterDF = pd.DataFrame(uniqueLevelsInStudentRoster)
uniqueLevelsInStudentRosterDF = uniqueLevelsInStudentRosterDF.rename(columns = {0: 'codeDescription'})
uniqueLevelsInStudentRosterDF = uniqueLevelsInStudentRosterDF.assign(codeCode=pd.Series(['GK','G1','G2','G3','G4','G5','G6',
                                                                                         'G7','G8','G9','G10','G11','G12',
                                                                                         'G4','GK']).values)
print('uniqueLevelsInStudentRosterDF: ')
print(uniqueLevelsInStudentRosterDF)

levelsLookup = uniqueLevelsInStudentRosterDF.set_index('codeDescription').to_dict()['codeCode']
print('levelsLookup:')
print(levelsLookup)

# Get unique survey years and gender for observation
uniqueSurveyYears = rawRosterDF['SchoolYear'].unique()
print("uniqueSurveyYears: ",)
print(uniqueSurveyYears)

uniqueGender = rawRosterDF['Gender'].unique()
print("uniqueGender: ")
print(uniqueGender)

uniqueRepeat = rawRosterDF['Repeat Previous Year Grade'].unique()
print("uniqueRepeat: ")
print(uniqueRepeat)

uniqueTrin = rawRosterDF['Transferred From which school'].unique()
print("uniqueTrin: ")
print(uniqueTrin)

uniqueTrout = rawRosterDF['Transferred TO which school'].unique()
print("uniqueTrout: ")
print(uniqueTrout)

uniqueDropout = rawRosterDF['Drop-Out'].unique()
print("uniqueDropout: ")
print(uniqueDropout)


uniqueSchoolsInStudentRoster = rawRosterDF['School Name'].unique()
uniqueSchoolsInStudentRosterDF = pd.DataFrame(uniqueSchoolsInStudentRoster)
uniqueSchoolsInStudentRosterDF = uniqueSchoolsInStudentRosterDF.rename(columns = {0: 'schName'})
uniqueSchoolsInStudentRosterDF

cleanMergedSchoolsRightJoinDF = pd.merge(schoolsDF, uniqueSchoolsInStudentRosterDF, on='schName', how='right')
cleanMergedSchoolsLeftJoinDF = pd.merge(schoolsDF, uniqueSchoolsInStudentRosterDF, on='schName', how='left')
cleanMergedSchoolsRightJoinDF


# Writing data to sheets
# writer = pd.ExcelWriter(outEnrolmentTransitData)
# writer.save()
# schoolsDF.to_excel(writer, sheet_name='Schools Lookups', index=False)
# cleanMergedSchoolsRightJoinDF.to_excel(writer, sheet_name='cleanMergedSchoolsLeftJoinDF', index=False)
# #cleanMergedSchoolsLeftJoinDF.to_excel(writer, sheet_name='cleanMergedSchoolsRightJoinDF', index=False)
# writer.save()

# Some pre-processing data cleanup

# Clean repeater, transfers in/out, dropouts
repeatLookup = {
    'No': 'No',
    'Yes': 'Yes',
    'NO': 'No',
    'YES': 'Yes',
    'yes': 'Yes',
    'Missing': 'No',
    ' ': 'No'
}

repeat = rawRosterDF['Repeat Previous Year Grade'].map(repeatLookup)
#trin = rawRosterDF['Transferred From which school'].map(trinLookup)
#trout = rawRosterDF['Transferred TO which school'].map(troutLookup)
#dropout = rawRosterDF['Drop-Out'].map(dropoutLookup)
rawRosterDF = rawRosterDF.assign(repeat=repeat)

# Clean grade levels and age
rawRosterDF = rawRosterDF.replace(to_replace={'Grade Level':levelsLookup})
rawRosterDF = rawRosterDF.rename(columns = {'Grade Level': 'enLevel'})
rawRosterDF = rawRosterDF.rename(columns = {'Age as of September 30 of that School Year': 'enAge'})

# Clean schools
rawRosterDF = rawRosterDF.replace(to_replace={'School Name':schoolsLookup})
closedSchools = ['Kanifay ECE Center', 'Colonia ECE Center', 'Mizpah Christian High School', 'Mizpah High', 'Rumung Elementary School',
                 'Nukaf Elem/Sapota Paata Elem']
rawRosterDF = rawRosterDF[~rawRosterDF['School Name'].isin(closedSchools)]
rawRosterDF = rawRosterDF.rename(columns = {'School Name': 'schNo'})
rawRosterDF = rawRosterDF.rename(columns = {'Full Name': 'Name'})
rawRosterDF = rawRosterDF.rename(columns = {'Date of Birth': 'DoB'})

# Remove student with unknown DoB and Age
rawRosterDF = rawRosterDF[~(rawRosterDF['DoB'] == 'Unknown')]

# Clean survey years and genders
surveyYearsLookup = {
    'SY2016-2017': 2016,
    'SY2015-2016': 2015,
    'SY2014-2015': 2014,
    'SY2013-2014': 2013,
    'SY2012-2013': 2012
}
gendersLookup = {
    'male': 'M',
    'female': 'F',
    'MAle': 'M',
    'MALE': 'M',
    'Female': 'F',
    'Male': 'M',
}
rawRosterDF = rawRosterDF.replace(to_replace={'SchoolYear':surveyYearsLookup, 'Gender':gendersLookup})
rawRosterDF = rawRosterDF.rename(columns = {'SchoolYear': 'svyYear'})
print('rawRosterDF:')
print(rawRosterDF.head(10))

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
# %%time
################################################################################
# Process enrollment data
# NOTE: Long to process: executed in 13m 48s on MacBook Pro, 2.9 GHz Intel Core i7, 16 GB 2133 MHz LPDDR3)
# (~7min 1s on iMac with i9 CPU and 32GB RAM)
################################################################################

outEnrolmentTransitData = os.path.join(cwd,'data/FSM/fsm-enrollment-initial-transit-data.xlsx')

# svyYear, schNo, Name, Gender, DoB, enAge, enLevel  
enrolmentColumns = ['svyYear','schNo','enAge','enLevel','enM','enF']
rawEnrolmentDF = pd.DataFrame(columns=enrolmentColumns)
rawEnrolmentDF

# df = rawRosterDF[(rawRosterDF['svyYear'] == 2016) &
#                         (rawRosterDF['schNo'] == 'CHK175') &
#                         (rawRosterDF['enAge'] == 6) &
#                         (rawRosterDF['enLevel'] == 'GK')]

# Work on small sample to get this working first
rawRosterCleanedSampleDF = rawRosterDF #[:1000]

rawEnrolmentDF = rawRosterCleanedSampleDF.drop_duplicates(subset=['svyYear','schNo','enAge','enLevel'])
rawEnrolmentDF = rawEnrolmentDF.drop(['Name','Gender','DoB'], 1)
rawEnrolmentDF = rawEnrolmentDF.assign(enM=0,enF=0)
rawEnrolmentDF.reset_index(drop=True, inplace=True)

print("rawEnrolmentDF Empty: ")
print(rawEnrolmentDF)
print("rawRosterCleanedSampleDF: ")
print(rawRosterCleanedSampleDF)

# # Check if record exist and if not create it.
for student in rawRosterCleanedSampleDF.itertuples(): # .iterrows():
    #print(student)

    # What enrolment record is this student to update
    enrolRecord = rawEnrolmentDF[(rawEnrolmentDF['svyYear'] == student.svyYear) &
                              (rawEnrolmentDF['schNo'] == student.schNo) &
                              (rawEnrolmentDF['enAge'] == student.enAge) &
                              (rawEnrolmentDF['enLevel'] == student.enLevel)]

    try: 
        if student.Gender == 'M':
            rawEnrolmentDF.iloc[enrolRecord.index, rawEnrolmentDF.columns.get_loc('enM')] = rawEnrolmentDF.iloc[enrolRecord.index, rawEnrolmentDF.columns.get_loc('enM')] + 1
        elif student.Gender == 'F':
            rawEnrolmentDF.iloc[enrolRecord.index, rawEnrolmentDF.columns.get_loc('enF')] = rawEnrolmentDF.iloc[enrolRecord.index, rawEnrolmentDF.columns.get_loc('enF')] + 1
    except IndexError:
        print(rawEnrolmentDF)
        print("Index at fault: ", enrolRecord.index)

#rawEnrolmentDF = pd.read_excel(outEnrolmentTransitData, sheet_name="EnrolmentsRaw", header=0, usecols="A:F")
rawEnrolmentDF = rawEnrolmentDF[(rawEnrolmentDF['enAge'] != -195) &
                                (rawEnrolmentDF['enAge'] != -1) &
                                (rawEnrolmentDF['enAge'] != 0) &
                                (~rawEnrolmentDF['enAge'].isnull())]
print("Uniques age values: ", rawEnrolmentDF['enAge'].unique())

# Drop some data which I may need to come back to later
rawEnrolmentDF = rawEnrolmentDF.drop(['Repeat Previous Year Grade','Transferred From which school','Transferred TO which school','Drop-Out','repeat'], 1)

# Writing enrolment data to sheets
writer = pd.ExcelWriter(outEnrolmentTransitData)
rawEnrolmentDF.to_excel(writer, sheet_name='EnrolmentsRaw', index=False)
writer.save()

print("rawEnrolmentDF: ")
print(rawEnrolmentDF)

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
################################################################################
# Process School Surveys
################################################################################

outInitialSchoolSurveysData = os.path.join(cwd,'data/FSM/fsm-schoolsurveys-initial-data.xlsx')

schoolTypesLookup = schoolsDF.set_index('schNo').to_dict()['schType']

rawEnrolmentDF = pd.read_excel(outEnrolmentTransitData, sheet_name="EnrolmentsRaw", header=0, usecols="A:F")
schoolSurveyDF = rawEnrolmentDF.drop(['enAge','enLevel'], 1)
schoolSurveyDF = schoolSurveyDF.groupby(['svyYear','schNo'], as_index=False).sum()
schoolSurveyDF = schoolSurveyDF.rename(columns = {'enF': 'ssEnrolF', 'enM': 'ssEnrolM'})
schoolSurveyDF = schoolSurveyDF.assign(ssEnrol = lambda x: x.ssEnrolF + x.ssEnrolM)
schoolSurveyDF.insert(0,'ssID',range(1,len(schoolSurveyDF.index)+1))

schoolSurveyDF['ssSchType'] = schoolSurveyDF['schNo'].map(schoolTypesLookup)
print("schoolSurveyDF: ")
print(schoolSurveyDF.head(3))

# Writing SchoolSurveys data to sheets
writer = pd.ExcelWriter(outInitialSchoolSurveysData)
schoolSurveyDF.to_excel(writer, sheet_name='SchoolSurveys', index=False)
writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
################################################################################
# Process final Enrolments ['ssID', 'enAge', 'enLevel', 'enM', 'enF']
################################################################################

outEnrolmentData = os.path.join(cwd,'data/FSM/fsm-enrollment-initial-data.xlsx')

schoolSurveyDF['ssIDTemp'] = schoolSurveyDF['svyYear'].map(str) + schoolSurveyDF['schNo']
schoolSurveyLookup = schoolSurveyDF.set_index(['ssIDTemp']).to_dict()['ssID']
print("schoolSurveyLookup: ")
print(schoolSurveyLookup)


enrolmentsDF = rawEnrolmentDF
enrolmentsDF['ssIDTemp'] = enrolmentsDF['svyYear'].map(str) + enrolmentsDF['schNo']
enrolmentsDF['ssID'] = enrolmentsDF['ssIDTemp'].map(schoolSurveyLookup)
enrolmentsDF = enrolmentsDF.drop(['svyYear','schNo','ssIDTemp'], 1)

order_cols = ['ssID', 'enAge', 'enLevel', 'enM', 'enF']
enrolmentsDF = enrolmentsDF[order_cols]

print("rawEnrolment: ")
print(rawEnrolmentDF.head(3))
print("schoolSurveyDF: ")
print(schoolSurveyDF.head(3))
print("enrolments: ")
print(enrolmentsDF)

# Writing Enrolments data to sheets
writer = pd.ExcelWriter(outEnrolmentData)
enrolmentsDF.to_excel(writer, sheet_name='Enrolments', index=False)
writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
################################################################################
# Process repeaters
################################################################################

outPupilTableData = os.path.join(cwd,'data/FSM/fsm-enrollment-pupil-table-initial-data.xlsx')

# Get all raw repeat data
rawRepeatDF = rawRosterDF[rawRosterDF['repeat'] == 'Yes']
rawRepeatDF = rawRepeatDF.drop(['Name','DoB','Repeat Previous Year Grade','Transferred From which school','Transferred TO which school','Drop-Out'],1)
# Add ssID using schoolSurveyLookup
rawRepeatDF['ssIDTemp'] = rawRepeatDF['svyYear'].map(str) + rawRepeatDF['schNo']
rawRepeatDF['ssID'] = rawRepeatDF['ssIDTemp'].map(schoolSurveyLookup)
print("rawRepeatDF: ")
print(rawRepeatDF.head(3))

# Prepare a pupiltable dataframe prefilled with all known values
# 'ssID','ptCode','ptAge','ptLevel','ptPage','ptRow','ptCol','ptM','ptF','ptSum','ptTableDef','ptTable'
# Note: only the necessary columns are included
pupilTableColumns = ['ssID','ptCode','ptAge','ptLevel','ptM','ptF']
pupilTableDF = pd.DataFrame(columns=pupilTableColumns)
pupilTableDF = rawRepeatDF.drop_duplicates(subset=['svyYear','schNo','enAge','enLevel'])
pupilTableDF = pupilTableDF.assign(ptM=0,ptF=0,ptCode='REP')
pupilTableDF.reset_index(drop=True, inplace=True)

print("pupilTableDF Empty: ")
print(pupilTableDF.head(3))

# # Test grabbing one of the repeater record
# df = rawRosterDF[(rawRosterDF['svyYear'] == 2016) &
#                  (rawRosterDF['schNo'] == 'CHK002') &
#                  (rawRosterDF['enAge'] == 7) &
#                  (rawRosterDF['enLevel'] == 'G1')]
# print("df: ")
# print(df)

# Fill up the dataframe with pupiltable data
for repeater in rawRepeatDF.itertuples(): # .iterrows():
    #print(repeater)

    # What pupiltable record is this repeater to update
    pupilTableRecord = pupilTableDF[(pupilTableDF['svyYear'] == repeater.svyYear) &
                                    (pupilTableDF['schNo'] == repeater.schNo) &
                                    (pupilTableDF['enAge'] == repeater.enAge) &
                                    (pupilTableDF['enLevel'] == repeater.enLevel)]

    try: 
        if repeater.Gender == 'M':
            pupilTableDF.iloc[pupilTableRecord.index, pupilTableDF.columns.get_loc('ptM')] = pupilTableDF.iloc[pupilTableRecord.index, pupilTableDF.columns.get_loc('ptM')] + 1
        elif repeater.Gender == 'F':
            pupilTableDF.iloc[pupilTableRecord.index, pupilTableDF.columns.get_loc('ptF')] = pupilTableDF.iloc[pupilTableRecord.index, pupilTableDF.columns.get_loc('ptF')] + 1
    except IndexError:
        print(pupilTableDF)
        print("Index at fault: ", pupilTableRecord.index)

# Make more consistent with pupiltable
pupilTableDF = pupilTableDF.drop(['svyYear','schNo','Gender','repeat','ssIDTemp'], 1)
pupilTableDF = pupilTableDF.rename(columns = {'enAge': 'ptAge', 'enLevel': 'ptLevel'})
        
print("pupilTableDF: ")
print(pupilTableDF.head(3))

# # Writing data to sheets
writer = pd.ExcelWriter(outPupilTableData)
pupilTableDF.to_excel(writer, sheet_name='PupilTableRepeaters', index=False)
writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
################################################################################
# Process accreditation data
################################################################################

rawAccreditationData = os.path.join(cwd,'data/FSM/AccreditationData_2017.xlsx')
outAccreditationData = os.path.join(cwd,'data/FSM/fsm-schools-accreditation-initial-data.xlsx')

rawAccreditationDataDF = pd.read_excel(rawAccreditationData, sheet_name="UpdatedSheet", header=1)

# Remove all records where no school inspection exists for now
rawAccreditationDataDF = rawAccreditationDataDF[~rawAccreditationDataDF['L1'].isnull()]

rawAccreditationDataDF['Year'].fillna(value=2016,inplace=True)

schoolWithExistingLookup = list(schoolsLookup.keys())
schoolWithExistingLookup.sort()
rawUniqueSchoolNamesFromAccreditation = list(rawAccreditationDataDF['School Name'].unique())
rawUniqueSchoolNamesFromAccreditation.pop()
rawUniqueSchoolNamesFromAccreditation.sort()
print('schoolWithExistingLookup: ', schoolWithExistingLookup[:3])
print('rawUniqueSchoolNamesFromAccreditation: ', rawUniqueSchoolNamesFromAccreditation[:3])
# set(rawUniqueSchoolNames).difference(schoolWithExistingLookup)

# No longer needed, I have a comprehensive mapping of school names with anomalies
# # First cleanup schools from raw accreditation data (to be fixed manually and re-entered into the schoolsLookup in other cell)
# uniqueSchoolsFromRawAccreditationDF = pd.DataFrame({'RawSchools': rawUniqueSchoolNamesFromAccreditation})
# #create a mapping for those schools that actually match a school from other records
# SchNoMapping = uniqueSchoolsFromRawAccreditationDF['RawSchools'].map(schoolsLookup)
# uniqueSchoolsFromRawAccreditationDF = uniqueSchoolsFromRawAccreditationDF.assign(SchNoMappingCompleteMe = SchNoMapping)
# SchNameMapping = uniqueSchoolsFromRawAccreditationDF['SchNoMappingCompleteMe'].map(schoolsLookupByName)
# uniqueSchoolsFromRawAccreditationDF = uniqueSchoolsFromRawAccreditationDF.assign(SchoolName = SchNameMapping)
# print('uniqueSchoolsFromRawAccreditationDF: ')
# print(uniqueSchoolsFromRawAccreditationDF.head(5))
# # Writing all schools anomalies for hand fixing
# # uniqueSchoolsFromRawAccreditationDF.to_excel(schoolTypoWriter, sheet_name='SchoolsFromAccreditation', index=False)
# # schoolTypoWriter.save()

# Prepare data for InspectionSet (inspsetID,inspsetName,inspsetType,inspsetYear)
InspectionSets = {
    'inspsetID': [1,2],
    'inspsetName': ['2016','2017'],
    'inspsetType': ['SCHACCR','SCHACCR'],
    'inspsetYear': [2016,2017]
}

inspectionSetsLookup = {
    2016: 1,
    2017: 2
}

inspectionSetDF = pd.DataFrame(data=InspectionSets)
print('inspectionSetDF:')
print(inspectionSetDF.head(5))
      
# Process SchoolAccreditation starting from the raw data
schNo = rawAccreditationDataDF['School Name'].map(schoolsLookup)
inspsetID = rawAccreditationDataDF['Year'].map(inspectionSetsLookup)
rawAccreditationDataDF = rawAccreditationDataDF.assign(schNo = schNo, inspsetID = inspsetID)
rawAccreditationDataDF = rawAccreditationDataDF.drop(['State','School Name','Column22','Column3','Column1','Year'], 1)
rawAccreditationDataDF = rawAccreditationDataDF.rename(columns = {'Date Visited': 'inspStart',
                                                                  'L1':'saL1', 'L2':'saL2', 'L3':'saL3', 'L4':'saL4',
                                                                  'T1':'saT1', 'T2':'saT2', 'T3':'saT3', 'T4':'saT4',
                                                                  'D1':'saD1', 'D2':'saD2', 'D3':'saD3', 'D4':'saD4',
                                                                  'N1':'saN1', 'N2':'saN2', 'N3':'saN3', 'N4':'saN4',
                                                                  'F1':'saF1', 'F2':'saF2', 'F3':'saF3', 'F4':'saF4',
                                                                  'S1':'saS1', 'S2':'saS2', 'S3':'saS3', 'S4':'saS4',
                                                                  'CO1':'saCO1', 'CO2':'saCO2',
                                                                  'Tally 1':'saLT1','Tally 2':'saLT2','Tally 3':'saLT3','Tally 4':'saLT4',
                                                                  'Total':'saT','Level':'saSchLevel'})
rawAccreditationDataDF.insert(0,'inspID',range(1,len(rawAccreditationDataDF.index)+1))
rawAccreditationDataDF.insert(0,'saID',rawAccreditationDataDF['inspID'])
# TODO remove records with no school ID for now until the schools are all sorted out
rawAccreditationDataDF = rawAccreditationDataDF[~rawAccreditationDataDF['schNo'].isnull()]

schoolAccreditationDF = rawAccreditationDataDF.drop(['inspID','inspStart','inspsetID','schNo'],1)
schoolAccreditationDF.fillna('NULL', inplace=True)
print('schoolAccreditationDF:')
print(schoolAccreditationDF.head(5))
      
# Process SchoolInspection data (inspID,schNo,inspPlanned,inspStart,inspEnd,inspNote,inspBy,inspsetID)
schoolInspectionDF = rawAccreditationDataDF.drop(list(schoolAccreditationDF.columns.values),1)
schoolInspectionDF.fillna('NULL', inplace=True)
print('schoolInspectionDF:')
print(schoolInspectionDF.head(5))
      
# Writing all school accreditation and related data to sheets
writer = pd.ExcelWriter(outAccreditationData)
inspectionSetDF.to_excel(writer, sheet_name='InspectionSet', index=False)
schoolAccreditationDF.to_excel(writer, sheet_name='SchoolAccreditation', index=False)
schoolInspectionDF.to_excel(writer, sheet_name='SchoolInspection', index=False)
writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
################################################################################
# Process teachers data
################################################################################

#rawTeachersData = os.path.join(cwd,'data/FSM/SY2015-2016CombinedTeachersending.xlsx')
rawTeachersData = os.path.join(cwd,'data/FSM/CombinedDataStaff.xlsx')
outTeachersData = os.path.join(cwd,'data/FSM/fsm-teachers-initial-data.xlsx')

# Prepare writer for all teacher related data
writer = pd.ExcelWriter(outTeachersData)

# All raw fields 'State','Municipality/Zone/Region','Island Name','First Name','Last Name','Job Title','Ethnicity','Citizenship','Staff Type','Teacher-Type','Organization','Gender','Highest Degree Achieved','Copy of Degree/Certificate Available','Field of Study','Certified','Expiration','Date of Hire ','Date of Birth','Annual Salary','Funding Source','School Name','School Type','School Level','Grade Taught','Employment Status','Reason','Date of Exit','Total # of days absent'
rawTeachersAppointmentsDF = pd.read_excel(rawTeachersData, sheet_name="CombinedSchoolStaff", header=0)
# Change all the string missing to actual pandas NULL since this is what we'll be inserting in DB
rawTeachersAppointmentsDF.replace('Missing', 'NULL', inplace=True)
# Handle as many dates as possible setting the bad ones to NaN 	 
tDatePSAppointed = pd.to_datetime(rawTeachersAppointmentsDF['Date of Hire '], errors="coerce").dt.date #, format="%m/%d/%Y"
tDOB = pd.to_datetime(rawTeachersAppointmentsDF['Date of Birth'], errors="coerce").dt.date #, format="%m/%d/%Y"
tDatePSClosed = pd.to_datetime(rawTeachersAppointmentsDF['Date of Exit'], errors="coerce").dt.date #, format="%m/%d/%Y"

# Only work on the first 500 rows until this is cleaned
#rawTeachersAppointmentsDF = rawTeachersAppointmentsDF #[0:499]
# Cleanup
# Drop what we will not need at all here onwards, rename, clean dates...
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF.assign(tDatePSAppointed = tDatePSAppointed, tDOB = tDOB, tDatePSClosed = tDatePSClosed)
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF.drop(['State','Municipality/Zone/Region','Island Name','School Type','School Level',
                                                            'Date of Hire ','Date of Birth','Date of Exit'], 1)
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF.rename(columns = {
    'First Name': 'tGiven', 'Last Name': 'tSurname', 'Gender': 'tSex', 'Reason': 'tCloseReason'
})
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF.fillna('NULL')

# Get a whole bunch of unique repeating values in lookup like columns
# as first step in cleaning them
uniqueStaffTypeFromRawTeacher = rawTeachersAppointmentsDF['Staff Type'].unique()
print('uniqueStaffTypeFromRawTeacher:')
print(uniqueStaffTypeFromRawTeacher)

uniqueTeacherTypeFromRawTeacher = rawTeachersAppointmentsDF['Teacher-Type'].unique()
print('uniqueTeacherTypeFromRawTeacher:')
print(uniqueTeacherTypeFromRawTeacher)

uniqueOrganizationFromRawTeacher = rawTeachersAppointmentsDF['Organization'].unique()
print('uniqueOrganizationFromRawTeacher:')
print(uniqueOrganizationFromRawTeacher)

uniqueCertifiedFromRawTeacher = rawTeachersAppointmentsDF['Certified'].unique()
print('uniqueCertifiedFromRawTeacher:')
print(uniqueCertifiedFromRawTeacher)

uniqueEthnicitiesFromRawTeacher = rawTeachersAppointmentsDF['Ethnicity'].unique()
print('uniqueEthnicitiesFromRawTeacher:')
print(uniqueEthnicitiesFromRawTeacher)

uniqueJobTitlesFromRawTeacher = rawTeachersAppointmentsDF['Job Title'].unique() # role in Pineapple
print('uniqueJobTitlesFromRawTeacher:')
print(uniqueJobTitlesFromRawTeacher)

uniqueHighestDegreesFromRawTeacher = rawTeachersAppointmentsDF['Highest Degree Achieved'].unique()
print('uniqueHighestDegreesFromRawTeacher:')
print(uniqueHighestDegreesFromRawTeacher) 

uniqueSchoolsFromRawTeacher = rawTeachersAppointmentsDF['School Name'].unique()
print('uniqueSchoolsFromRawTeacher:')
print(uniqueSchoolsFromRawTeacher)

uniqueTeacherAppointmentSchoolYearFromRawTeacher = rawTeachersAppointmentsDF['SchoolYear'].unique()
print('uniqueTeacherAppointmentSchoolYearFromRawTeacher:')
print(uniqueTeacherAppointmentSchoolYearFromRawTeacher)

uniqueGenderFromRawTeacher = rawTeachersAppointmentsDF['tSex'].unique()
print('uniqueGenderFromRawTeacher:')
print(uniqueGenderFromRawTeacher)

# Manually constructed lookups from above unique values
# these must be rebuilt when/if new data would come in as input
rawNationalitiesLookup = {
    'Chuukese': 'CHU',
    'American': 'USA',
    'Australian': 'AUS',
    'Romanian': 'ROU',
    'Finnish': 'FIN',
    'Belgian': 'BEL',
    'Yapese ': 'YAP',
    'Filipino': 'PHL',
    'Brazilian': 'BRA',
    'Russian': 'RUS',
    'Other': 'O',
    'Pohnpeian': 'PNI',
    'Indonesian': 'IDN',
    'Japanese': 'JPN',
    'Yap': 'YAP',
    'Chuuk': 'CHU',
    'Yapese': 'YAP',
    'Palauan': 'PLW',
    'Caucasian': 'USA',
    'N. American': 'USA',
    'Pakistani': 'PAK',
    'Vietnamese': 'VNM',
    'Norwegian': 'NOR' 
}

rawRolesLookup = {
    'Classroom Teacher I': 'CTI',
    'Classroom Teacher II': 'CTII',
    'Classroom Teacher IV': 'CTIV',
    'Classroom Teacher III': 'CTIII',
    'Classroom Teacher V': 'CTV',
    'Classroom Mentor': 'CM',
    'Vocational Coordinator': 'VC',
    'Classroom Teacher': 'CT',
    'School Principal II': 'SPII',
    'Vocational Teacher II': 'VTII',
    'Assistant Principal III': 'ASPIII',
    'School Principal I': 'SPI',
    'Teacher Assistant': 'TA',
    'Houseparent I': 'HI',
    'Cook III': 'CIII',
    'Head Teacher': 'HT',
    'School Principal III': 'CPIII',
    'Peace Corp Volunteer': 'PCV',
    'Acting Principal I': 'API',
    'School Principal (Middle School)': 'SP',
    'School Principal(Primary Grade)': 'SP',
    'School Principal': 'SP',
    'Classroom Teacher_Regular': 'CT',
    'Classroom Teacher_ECE': 'CT',
    'Classroom Teacher_Special Ed.': 'CT',
    'Teacher Aide_WD&ST': 'TA',
    'Principal/Classroom Teacher_Regular': 'SP',
    'Teacher Aide_Special Ed': 'TA',
    'Teacher Aide_Special Ed.': 'TA',
    'Classroom Teacher_VocEd': 'VTI',
    'Teacher Aide_ECE': 'TA',
    'Related Services Assistant': 'TA',
    'Clasroom Teacher_Regular': 'CT',
    'Principal': 'SP',
    'Teacger Aide_WD&ST': 'TA',
    'Principal III': 'SPIII',
    'Classroom Teacher': 'CT',    
    'Vice Principal I': 'SVPI',
    'Classroom Teacher I (SpEd)': 'CTI',
    'Classroom Teacher II (SpEd)': 'CTII',
    'Classroom Teacher (Contract)': 'CT',
    'School Principal II': 'SPII',
    'Assistant Principal ': 'AP',
    'Vocational Teacher': 'VT',
    'Vocational Teacher I': 'VTI',
    'Classroom Teacher III (SpEd)': 'CTIII',
    'Classroom Teacher  II': 'CTII',
    'Classroom Teacher  I': 'CTI',
    'Classroom Teache II': 'CTII',
    'Vocational Teacher III': 'VTIII',
    'Classroom Teacher IV (SpEd)': 'CTIV',    
    'Classroom Teacher (SpEd)': 'CT',
    'Principal I': 'SPI',
    'Calssroom Teacher II': 'CTII',
    'Acting School Principal': 'AP',
    'Calssroom Teacher I': 'CTI',
    'Classroom Teacher Ii': 'CTII',
    'Classroom Teacher  (SpEd)': 'CT',
    'Classroom Teacher II ': 'CTII',
    'Classroom Teacher (Contract)1yr': 'CT',
    'Classroom Teacher (Contract 1YR)': 'CT',
    'School Principal III': 'SPIII',
    'Classroom Teacher I (Contract-NTE 1YR)': 'CTI',
    'Acting Principal I': 'API', #test dup
    'Vocational Instructor': 'VTI',
    'Classroom Teacher i': 'CTI',
    'School Principal II (Contract)': 'SPII',
    'Acting Head Teacher': 'AHT',
    'Classroom Teacher l': 'CTI',
    'School Prncipal': 'SP',
    'School Princiapl II': 'SPII',
    'School Principal (Middle School)': 'SP',
    'School Principal (Primary Grade)': 'SP',
    'Teacher Aide': 'TA',
    'Vice Principal': 'SVP',
    'Classroom Teacher (PCV)': 'CT',
    'Classroom Teacher (Bible)': 'CT',
    'Culture Teacher': 'CT',
    'Vocational Education Teacher': 'VT',
    'Culture Resource Teacher': 'CT',
    'Teacher ': 'CT',    
    'Teacher': 'CT',
    'Librarian': 'L',
    'Cook': 'C',
    'Bus Driver': 'BD',
    'Clerk': 'CL',
    'Couselor': 'SC',
    'Assistant Marine Instructor': 'VT',
    'Principla': 'SP',
    'Counselor': 'SC',
    'Kitchen Helper': 'KH',
    'Classroom Teacher-Regular': 'CT',
    'Prinipal': 'SP',
    'Boat Operator': 'BO',
    'ClassroomTeacher_Regular': 'CT',
    'Houseparent': 'H',
    'Driver': 'D',
    'Supply Technician': 'ST',
    'Print Disability Specialist': 'PDS',
    'Secretary': 'SE',
    'House Parent': 'HP',    
    'Education Specialis': 'ES',
    'Pricipal': 'SP',
    'School PrincipalIII': 'SPIII',
    'Classroom Mentor': 'CM',    
    'Vocational Coordinator': 'VC',
    'Consultant (Reform Plan)': 'CO',
    'Counselor IV': 'SCIV',
    'Secretary I': 'SEI',
    'Maintenance': 'MA',
    'Security Guard I': 'SGI',
    'Security Guard Supervisor': 'SGS',
    'Campus Maintenance': 'MA',
    'Registrar': 'RE',
    'School PrincipalII': 'SPIII',
    'Registrar I': 'REI',    
    'Security Guard': 'SG',
    'Trademan': 'TR',
    'Clerk Typist III': 'CLIII',
    'Custodian': 'CU',
    'Clerk Typist': 'CL',
    'Building Maitenance I': 'MAI',
    'Security Guard II': 'SGII',
    'Cook III': 'CIII',
    'Houseparent I': 'HI',
    'Cook I': 'CI',
    'School PrincipalI': 'SPI',
    'Assistant School PrincipalIII': 'ASPIII',
    'Data Clerk I': 'CI',
    'Peace Corp Volunteer': 'PCV',
    'School Principal(Middle School)': 'SP',
    'School Principal(Primary Grade)': 'SP',
    'Primary Consulting Resource Teacher': 'C',
    'Substitute Teacher': 'SUB',
    'School Counselor': 'SC',
    'Secondary Consulting Resource Teacher': 'C',
    'Administrative Officer': 'AO',
    'Maintenance Specialist': 'MA',
    'Secondary Transition Specialist': 'ES',
    'Secondary Transition Supervisor': 'TS',
    'PE Instuctor': 'VT',
    'School Accountant/Administrative Assistant': 'AO',
    'Supervisor': 'SU',
    'Moonitor': 'MO',
    'Principal/Administrator': 'SP',
    'Monitor': 'MO',
    'Canteen supervisor': 'CS',
    'Admin. Assistant': 'AA',
    'Voc Ed. Coordinator': 'VC',
    'Vice Princiapl': 'SVP',
    'School Counselor V': 'SCV',
    'Maintenance Worker III': 'MAIII',
    'Security Guard III': 'SGIII',
    'Bus Driver II': 'BDII',
    'Bus Driver I': 'BDI',
    'Data Clerk IV': 'CLIV',
    'Custodial Worker II'
    'Registrar II': 'REII',
    'Teacher Assistant': 'TA',
    'Maintenance Worker I': 'MAI',    
    'House Parent II': 'HII',
    'Librarian I': 'LI',
    'School Counselor I': 'SCI',
    'Cook II': 'CII',
    'Assistant Principal': 'ASP',
    'Resource Teacher': 'CT',
    'Teacher Aide (WD&ST)': 'TA',
    'Cook ': 'C',    
    'Teaching Principal': 'SP',
    'Bus Driver ': 'BD',
    'Contract Instructor': 'VT',
    'Home Arts Instructor': 'VT',
    'cook': 'C',
    'Ground Keeper': 'GK',
    'Assistant Librarian': 'AL',
    'Job Career Counselor': 'SC',
    'Related Services Assstant': 'RSA',
    'ECE Supervisor': 'SU',
    'Resource Teacher ': 'CT',
    'Teacher-WDNST': 'CT',
    'Teacher/PE Instuctor': 'CT',
    'Asst Director': 'ADI',
    'Director': 'DI',
    'Accountant': 'ACC',
    'Consultant  (R/Plan)': 'C',
    'Classroom Teacer IV': 'CTIV',
    'Acting Principal': 'AP',
    'Classroom Teache I': 'CTI',
    'Acting School Principal I': 'API',
    'World Teach': 'U'
}

# staff, teacher type, organisation for the 
rawDegreesLookup = {
    'AS': 'AS',
    'BA': 'BA',
    'AA': 'AA',
    'MA': 'MA',
    'none': 'NULL',
    'BS': 'BS',
    'As': 'AS',
    'None': 'NULL',
    'AS ': 'AS',
    'NULL': 'NULL',
    'AAS': 'AAS',
    'MS': 'MS',
    'BA ': 'BA',
    'HS Graduate': 'HS',
    'AA/AS': 'AS',
    'No Degree': 'NULL',
    'BA/BS': 'BS',
    'MA/MS': 'MS',
    'Certificate of Achievement': 'NULL',
    'AS Degree': 'AS',
    'Other': 'NULL',
    'AAA': 'AAA',
    'High School Diploma': 'HS',
    'AA Degree': 'AA',
    'BS ': 'BS',
    'BS, Diploma': 'BS',
    'CA': 'C',
    'PHD': 'PHD',
    'BS/MA': 'MA',
    'Finished 8th Gr': 'NULL',
    'High School': 'HS',
    'Some High Schoo': 'NULL',
    'AA\\': 'AA',
    'Some Elementary': 'NULL',
    'Some High School': 'NULL',
    'BS/B.Sc.': 'BS',
    'M.Ed.': 'MED',
    '3RD YR.': 'NULL',
    'Certificate': 'C',
    'BBA': 'BA',
    'Elementary Graduate': 'NULL',
    'Third Year': 'NULL',
    'Certificate of Completion': 'NULL'
}

# In the EMIS, there is no specific place to stored the equivalent of the columns "Staff Type", "Teacher-Type", "Organization"
      
# Built based on Staff Type, Teacher-Type, Organization
#       Local Regular Teaching Staff
#       Local Special Education Teaching Staff
#       Local Early Childhood Teaching Staff
#       Local Volunteer Teaching Staff
#       DOE Regular Teaching Staff
#       DOE Special Education Teaching Staff
#       DOE Early Childhood Teaching Staff
#       DOE Volunteer Teaching Staff
#       World Teacher Regular Teaching Staff
#       World Teacher Special Education Teaching Staff
#       World Teacher Early Childhood Teaching Staff
#       World Teacher Volunteer Teaching Staff
#       Peace Corp Regular Teaching Staff
#       Peace Corp Special Education Teaching Staff
#       Peace Corp Early Childhood Teaching Staff
#       Peace Corp Volunteer Teaching Staff
      
staffTypeFromRawTeacherLookup = {
    'Teaching Staff': 'TS',
    'Teaching Staff-TS': 'TS',
    'None Teaching Staff': 'NTS',
    'TS': 'TS',
    'NTS': 'NTS',
    'Principal': 'TS',
    'Vice Principal': 'TS',
    'Counselor': 'NTS',
    'Admin. Assistant': 'NTS',
    'Clerk': 'NTS',
    'Voc Ed. Coordinator': 'TS',
    'Maintenance': 'NTS',
    'Librarian': 'NTS',
    'Clerk ': 'NTS',
    'Librarian Staff': 'NTS',
    'Clerk Staff': 'NTS',
    'Supervisor': 'NTS',
    'T': 'TS',
    'NT': 'NTS',
    'Volunteer': 'TS',
    'NST': 'NTS',
    'Library': 'NTS'
}

certifiedFromRawTeacherLookup = {
    'NULL': 'NULL',
    'Certified': 'NSTT',
    'Yes': 'NSTT',
    'No': 'NULL',
    'yes': 'NSTT',
    ' Yes': 'NSTT',
    'Processing': 'NULL'
}

genderFromRawTeacherLookup = {
    'NULL': 'NULL',
    'Male': 'M',
    'Female': 'F'
}

# Add helper column to identify teachers
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF.assign(theTeacher = lambda x: x.tGiven + '-' + x.tSurname)
# clean staff type into another column staffTypeTemp
staffTypeTemp = rawTeachersAppointmentsDF['Staff Type'].map(staffTypeFromRawTeacherLookup)
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF.assign(staffTypeTemp = staffTypeTemp)
# clean certifications into another column certifiedTemp    
certifiedTemp = rawTeachersAppointmentsDF['Certified'].map(certifiedFromRawTeacherLookup)
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF.assign(staffTypeTemp = staffTypeTemp, certifiedTemp = certifiedTemp)
rawTeachersAppointmentsDF.to_excel(writer, sheet_name='RawTeacherAppointments', index=False)
# remove all non-teachers for now as raw data contains all staff
rawTeachersAppointmentsDF = rawTeachersAppointmentsDF[rawTeachersAppointmentsDF['staffTypeTemp'] == 'TS']
# print('rawTeachersAppointmentsDF:')
# print(rawTeachersAppointmentsDF.head(3))
      
# Teacher DB (tID,tDOB,tSex,tGiven,tSurname)
teachersDF = rawTeachersAppointmentsDF[['theTeacher','tDOB','tSex','tGiven','tSurname',
                                        'Highest Degree Achieved','Field of Study','certifiedTemp',
                                        'tDatePSAppointed','tDatePSClosed','tCloseReason']]
teachersDF = teachersDF.replace({'tSex': {'Male': 'M', 'Female': 'F'}})
# Small possibility of loosing some data here (some teacher duplicate might contain the certification while others not)
# the drop_dup only keeps the first occurance
teachersDF = teachersDF.drop_duplicates(['theTeacher'])
teachersDF.insert(0,'tID',range(1,len(teachersDF.index)+1))
# TODO Fix single quote (') in teacher names
# create teachers lookup
teachersLookup = teachersDF.set_index('theTeacher').to_dict()['tID']
print('teachersDF:')
print(teachersDF.head(1))

# teacherTypeFromRawTeacherLookup = {
#     'Missing': Nan,
#     'Regular-R': 'R',
#     'Special Ed.': 'SE',
#     'ECE': 'ECE',
#     'R': 'R',
#     'SE': 'SE',
#     'Volunteer': 'V',
#     'RSA': ?,
#     'Regular': 'R',
#     'SPED': 'SE',
#     'Early Childhood Education': 'ECE',
#     'SDA': ?,
#     'World Teacher': ?,
#     'AVS II', ?,
#     'SM': ?,
#     'RT': 'R',
#     'SET': 'SE',
#     'Resident': ?,
#     'Local': ?,
#     'WT': ?
# }
      
# organizationFromRawTeacherLookup = {
#     'Missing': NaN,
#     'Local': '',
#     'CDOE': 'DOE',
#     'Special Education': 'DOE',
#     'World Teacher': 'WT',
#     'Peace Corp': 'PC',
#     'Private School': ?,
#     'JIV': ?,
#     'DOE': 'DOE',
#     'DOE ': 'DOE',
#     'KDOE': ?,
#     'World Teach': 'WT',
#     'KMG': ?,
#     'YCHS': ?
#     'Private': ?
# }

# No longer needed, I have a comprehensive mapping of school names with anomalies
# # First cleanup schools (to be written to file for manual association to correct school)
# uniqueSchoolsFromRawTeacherDF = pd.DataFrame({'RawSchools': uniqueSchoolsFromRawTeacher})
# #create a mapping for those schools that actually match a school from other records
# SchNoMapping = uniqueSchoolsFromRawTeacherDF['RawSchools'].map(schoolsLookup)
# uniqueSchoolsFromRawTeacherDF = uniqueSchoolsFromRawTeacherDF.assign(SchNoMappingCompleteMe = SchNoMapping)
# SchNameMapping = uniqueSchoolsFromRawTeacherDF['SchNoMappingCompleteMe'].map(schoolsLookupByName)
# uniqueSchoolsFromRawTeacherDF = uniqueSchoolsFromRawTeacherDF.assign(SchoolName = SchNameMapping)
# # Writing all schools anomalies for hand fixing
# schoolTypoWriter = pd.ExcelWriter(outInitialSchoolsTyposData)
# uniqueSchoolsFromRawTeacherDF.to_excel(schoolTypoWriter, sheet_name='SchoolsFromTeachers', index=False)
# schoolTypoWriter.save()
      
# Process teacher training (tID,trInstitution,trQual,trMajor) or (tID,trQual,trMajor)
# this is the academic degrees
teachersTrainingDF = teachersDF[['tID','Highest Degree Achieved','Field of Study']]
trQual= teachersTrainingDF['Highest Degree Achieved'].map(rawDegreesLookup) #teachersTrainingDF.loc['trQual']
teachersTrainingDF = teachersTrainingDF.rename(columns = {'Field of Study': 'trMajor'})
teachersTrainingDF = teachersTrainingDF.assign(trQual = trQual)
teachersTrainingDF = teachersTrainingDF.drop(['Highest Degree Achieved'], 1)
# and now the NSTT certification
teachersCertifiedDF = teachersDF[['tID','certifiedTemp']]
teachersCertifiedDF = teachersCertifiedDF[teachersCertifiedDF['certifiedTemp'] == 'NSTT']
teachersCertifiedDF = teachersCertifiedDF.rename(columns = {'certifiedTemp': 'trQual'})
#teachersCertifiedDF.insert(0,'trInstitution','FSM National Standard Teacher Certification') # range(1,len(teachersCertifiedDF.index)+1)
teachersTrainingDF = teachersTrainingDF.append(teachersCertifiedDF)
teachersTrainingDF = teachersTrainingDF.fillna('NULL')
teachersTrainingDF = teachersTrainingDF[~(teachersTrainingDF['trQual'] == 'NULL')]
print(teachersTrainingDF.head(1))

# Process teacher appointments. (tID,taDate,SchNo,taRole,estpNo,taEndDate)
teachersAppointmentsDF = rawTeachersAppointmentsDF[['theTeacher','SchoolYear','School Name','Job Title']]
# Set taDate and taEndDate based on SchoolYear ??
teachersAppointmentStartDatesLookup = {
    'SY2013-2014': '2013-08-01 00:00:00.000',
    'SY2014-2015': '2014-08-01 00:00:00.000',
    'SY2015-2016': '2015-08-01 00:00:00.000',
    'SY2016-2017': '2016-08-01 00:00:00.000',
}
teachersAppointmentEndDatesLookup = {
    'SY2013-2014': '2014-07-31 00:00:00.000',
    'SY2014-2015': '2015-07-31 00:00:00.000',
    'SY2015-2016': '2016-07-31 00:00:00.000',
    'SY2016-2017': '2017-07-31 00:00:00.000',
}
taDate = teachersAppointmentsDF['SchoolYear'].map(teachersAppointmentStartDatesLookup)
taEndDate = teachersAppointmentsDF['SchoolYear'].map(teachersAppointmentEndDatesLookup)
teachersAppointmentsDF = teachersAppointmentsDF.assign(taDate = taDate, taEndDate = taEndDate)
# # Need to link to correct teacher
tID = teachersAppointmentsDF['theTeacher'].map(teachersLookup)
# Need to link to correct schools
SchNo = teachersAppointmentsDF['School Name'].map(schoolsLookup)
taRole = teachersAppointmentsDF['Job Title'].map(rawRolesLookup)
teachersAppointmentsDF = teachersAppointmentsDF.assign(tID = tID, SchNo = SchNo, taRole = taRole)
# remove all appointments where we don't know the school and role since this is invalid data
# by nature (how can this be an appointment without knowing the school/role)
teachersAppointmentsDF = teachersAppointmentsDF[(pd.notnull(teachersAppointmentsDF['SchNo'])) & (pd.notnull(teachersAppointmentsDF['taRole']))]
print('teachersAppointmentsDF:')
print(teachersAppointmentsDF.head(1))

# Process establishments (estpNo,schNo,estpRoleGrade,estpActiveDate,estpTitle)
estpPositions = {}

def getEstpNo(row):
    estpKey = str(row.schNo) + '-' + str(row.taRole)
    if (estpKey in estpPositions):
        # Next position to assign to the school
        schoolPositions = estpPositions[estpKey]
        lastPositionIndex = schoolPositions.pop()
        newPositionIndex = lastPositionIndex+1
        estpPositions[estpKey] = schoolPositions + [lastPositionIndex,newPositionIndex]
        schoolRole = estpKey + '-' + str(newPositionIndex)
        return schoolRole
    else:
        # First position assigned to the school
        estpPositions[estpKey] = [1]
        schoolRole = str(estpKey) + '-1'        
        return schoolRole
    
# To simplify things we'll create a new establishment school position for each teacher appointment
# This is a simpler way to get off the ground faster, I think (hope)
#teachersAppointmentsWithEstpSetDF = teachersAppointmentsDF[teachersAppointmentsDF['taDate'] == '2016-08-01 00:00:00.000']
#establishmentsDF = teachersAppointmentsDF[teachersAppointmentsDF['taDate'] == '2016-08-01 00:00:00.000']
establishmentsDF = teachersAppointmentsDF
establishmentsDF = establishmentsDF.rename(columns = {'taDate': 'estpActiveDate', 'SchNo': 'schNo'})
estpTitle = establishmentsDF['taRole'].map(teacherRoleByNameLookups)
estpRoleGrade = establishmentsDF['taRole'].map(roleGradesLookups)
establishmentsDF = establishmentsDF.assign(estpTitle = estpTitle, estpRoleGrade = estpRoleGrade)
estpNo = establishmentsDF.apply(lambda row: getEstpNo(row),axis=1)
establishmentsDF.insert(0, 'estpNo', estpNo)
teachersAppointmentsDF = teachersAppointmentsDF.assign(estpNo = estpNo)
print('establishmentsDF:')
print(establishmentsDF.head(1))

# final cleanups of unecessary columns
teachersDF = teachersDF.drop(['theTeacher','Highest Degree Achieved','Field of Study','certifiedTemp'], 1)
teachersAppointmentsDF = teachersAppointmentsDF.drop(['SchoolYear','School Name','Job Title','theTeacher'], 1)
establishmentsDF = establishmentsDF.drop(['tID','taRole','taEndDate','theTeacher','SchoolYear','School Name','Job Title'], 1)

# Write sheets
teachersDF.to_excel(writer, sheet_name='Teacher', index=False)
teachersTrainingDF.to_excel(writer, sheet_name='TeacherTraining', index=False)
teachersAppointmentsDF.to_excel(writer, sheet_name='TeacherAppointment', index=False)
establishmentsDF.to_excel(writer, sheet_name='Establishment', index=False)
writer.save()

# %% ein.hycell=false ein.tags="worksheet-0" jupyter={"outputs_hidden": false} slideshow={"slide_type": "-"}
import os
import xml.etree.ElementTree as ET
import pandas as pd

from xml.etree.ElementTree import Element

vermpaf = os.path.join(cwd,'data/FSM/VERMPAF-FEDEMIS-TEST-11mar2018.xml')

tree = ET.parse(vermpaf)
root = tree.getroot()

# Create unique list of  in exams sample
#survivals = set([]) # to insert tuples of schema (School_ID, School_Name)

# Process nation flows (i.e. transition rate, survival rate)
      # <TRM>
      # <TRF>
      # <TR>
      # <SRM>
      # <SRF>
      # <SR>
      # year="2017" yearOfEd="11"
data = []      
for s in root.iter('Survival'):
    if s.find('TRM') is not None:
       TRM = s.find('TRM').text
    else:
       TRM = None
    if s.find('TRF') is not None:
       TRF = s.find('TRF').text
    else:
       TRF = None
    if s.find('TR') is not None:
       TR = s.find('TR').text
    else:
       TR = None
    if s.find('SRM') is not None:
       SRM = s.find('SRM').text
    else:
       SRM = None
    if s.find('SRF') is not None:
       SRF = s.find('SRF').text
    else:
       SRF = None
    if s.find('SR') is not None:
       SR = s.find('SR').text
    else:
       SR = None       
    data.append({'TRM': TRM, 'TRF': TRF, 'TR': TR, 'SRM': SRM, 'SRF': SRF, 'SR': SR,
                 'year': s.get('year'), 'yearOfEd': s.get('yearOfEd')})

data2 = []      
for t in root.iter('TeacherQC'):

    if t.find('Enrol') is not None:
       Enrol = t.find('Enrol').text
    else:
       Enrol = None
    if t.find('TeachersM') is not None:
       TeachersM = t.find('TeachersM').text
    else:
       TeachersM = None
    if t.find('TeachersF') is not None:
       TeachersF = t.find('TeachersF').text
    else:
       TeachersF = None
    if t.find('Teachers') is not None:
       Teachers = t.find('Teachers').text
    else:
       Teachers = None
    if t.find('QualM') is not None:
       QualM = t.find('QualM').text
    else:
       QualM = None
    if t.find('QualF') is not None:
       QualF = t.find('QualF').text
    else:
       QualF = None
    if t.find('Qual') is not None:
       Qual = t.find('Qual').text
    else:
       Qual = None
    if t.find('CertM') is not None:
       CertM = t.find('CertM').text
    else:
       CertM = None
    if t.find('CertF') is not None:
       CertF = t.find('CertF').text
    else:
       CertF = None
    if t.find('Cert') is not None:
       Cert = t.find('Cert').text
    else:
       Cert = None
    if t.find('CertPercM') is not None:
       CertPercM = t.find('CertPercM').text
    else:
       CertPercM = None
    if t.find('CertPercF') is not None:
       CertPercF = t.find('CertPercF').text
    else:
       CertPercF = None
    if t.find('CertPerc') is not None:
       CertPerc = t.find('CertPerc').text
    else:
       CertPerc = None
    if t.find('QualPercM') is not None:
       QualPercM = t.find('QualPercM').text
    else:
       QualPercM = None
    if t.find('QualPercF') is not None:
       QualPercF = t.find('QualPercF').text
    else:
       QualPercF = None
    if t.find('QualPerc') is not None:
       QualPerc = t.find('QualPerc').text
    else:
       QualPerc = None
    if t.find('PTR') is not None:
       PTR = t.find('PTR').text
    else:
       PTR = None
    if t.find('CertPTR') is not None:
       CertPTR = t.find('CertPTR').text
    else:
       CertPTR = None
    if t.find('QualPTR') is not None:
       QualPTR = t.find('QualPTR').text
    else:
       QualPTR = None

    data2.append({'Enrol': Enrol, 'TeachersM': TeachersM, 'TeachersF': TeachersF, 'Teachers': Teachers,
                  'QualM': QualM, 'QualF': QualF, 'Qual': Qual, 'CertM': CertM, 'CertF': CertF, 'Cert': Cert, 
                  'CertPercM': CertPercM, 'CertPercF': CertPercF, 'CertPerc': CertPerc,
                  'QualPercM': QualPercM, 'QualPercF': QualPercF, 'QualPerc': QualPerc, 
                  'PTR': PTR, 'CertPTR': CertPTR, 'QualPTR': QualPTR, 
                  'year': t.get('year'), 'sectorCode': t.get('sectorCode')})

flowDF = pd.DataFrame(data)
teacherQualDF = pd.DataFrame(data2)
teacherQualDF = teacherQualDF.sort_values(by=['year', 'sectorCode'])

flowDF
teacherQualDF

# Writing data to sheets
outData = os.path.join(cwd,'data/FSM/fsm-flow-and-teacher-qual-data.xlsx')
writer = pd.ExcelWriter(outData)
flowDF.to_excel(writer, sheet_name='FlowData', index=False)
teacherQualDF.to_excel(writer, sheet_name='TeacherQualData', index=False)
writer.save()

# %%
