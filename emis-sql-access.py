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

# %%
# Import configuration
import json
from datetime import datetime
import os
import pandas as pd
import pickle

def load_config(config_path="config.json"):
    """Load configuration from a JSON file."""
    with open(config_path, 'r') as file:
        config = json.load(file)
    return config["sqlserver_name"], config["sqlserver_db"], config["sqlserver_ip"], config["sqlserver_port"], config["sqlserver_user"], config["sqlserver_pwd"], config['base_url'], config['username'], config['password'], config['output_directory'], config['source_workbook_filename']
    

# Test loading configuration
sqlserver_name, sqlserver_db, sqlserver_ip, sqlserver_port, sqlserver_user, sqlserver_pwd, base_url, username, password, output_directory, source_workbook_filename = load_config()
print("Configuration loaded successfully.")

# %%
from sqlalchemy import create_engine
import urllib

# Build the connection string for SQLAlchemy using pyodbc
params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={sqlserver_ip},{sqlserver_port};"
    f"DATABASE={sqlserver_db};"
    f"UID={sqlserver_user};"
    f"PWD={sqlserver_pwd};"
    f"TrustServerCertificate=yes;"
)

# Create SQLAlchemy engine
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")


# %%
import pandas as pd

# Define your SQL query
query = """
SELECT 
    S.stuCardID,
    S.stuGiven,
    S.stuMiddleNames,
    S.stuFamilyName,
    S.stuGender,
    S.stuDoB,
    S.stuDoBEst,
    S.stuEthnicity,
    S.stuMaritalStatus,
    S.stuSpEdIntakeDate,
	CASE 
        WHEN SE.stueSpEd = 1 THEN 'Yes'
        WHEN SE.stueSpEd = 0 THEN 'No'
        ELSE NULL
    END AS stueSpEdStr,
    CASE 
        WHEN SE.stueSpEdIEP = 1 THEN 'Yes'
        WHEN SE.stueSpEdIEP = 0 THEN 'No'
        ELSE NULL
    END AS stueSpEdIEPStr,
	CASE 
        WHEN SE.stueSpEdHasAccommodation = 1 THEN 'Yes'
        WHEN SE.stueSpEdHasAccommodation = 0 THEN 'No'
        ELSE NULL
    END AS stueSpEdHasAccomodationStr,
	SE.*,
	ENV.codeDescription SpEdEnv,
	DIS.codeDescription SpEdDis,
	ENG.codeDescription SpEdEng,
	ACC.codeDescription SpEdAcc,
	ASS.codeDescription SpEdAss
FROM [dbo].[StudentEnrolment_] SE
INNER JOIN Student_ S ON SE.stuID = S.stuID
LEFT OUTER JOIN lkpSpEdEnvironment ENV ON SE.stueSpEdEnv = ENV.codeCode
LEFT OUTER JOIN lkpDisabilities DIS ON SE.stueSpEdEnv = DIS.codeCode
LEFT OUTER JOIN lkpEnglishLearner ENG ON SE.stueSpEdEnglish = ENG.codeCode
LEFT OUTER JOIN lkpSpEdAccommodations ACC ON SE.stueSpEdAccommodation = ACC.codeCode
LEFT OUTER JOIN lkpSpEdAssessmentTypes ASS ON SE.stueSpEdAssessment = ASS.codeCode
ORDER BY stueYear
"""

# Run the query and get the result in a DataFrame
df_enrolments = pd.read_sql_query(query, engine)
# %store df_enrolments

# Preview the result
df_enrolments.head()

# %%
# Define your SQL query
query = """
WITH RankedTeachers AS (
    SELECT 
        TI.[tID],
        [tRegister],
        [tPayroll],
        [tDOB],
        [tDOBEst],
        [tSex],
        [tGiven],
        [tMiddleNames],
        [tSurname],
        [tGivenSoundex],
        [tSurnameSoundex],
        [tDatePSAppointed],
        [tDatePSClosed],
        [tCloseReason],
        [tchsID],
        TS.[ssID],
        [tchSalary],
        [tchCitizenship],
        [tchSponsor],
        [tchEdQual] AS [EdQualCode],
        TQ.codeDescription AS [EdQual],
        [tchQual] AS [QualCode],
        TQ2.codeDescription AS [Qual],
        [tchRole],
        [tchTAM],
        TS.[tID] AS TS_tID,
        SS.schNo,
        SS.svyYear,
        ROW_NUMBER() OVER (PARTITION BY TI.tID ORDER BY SS.svyYear DESC) AS rn
    FROM [dbo].[TeacherIdentity] TI
    INNER JOIN [dbo].[TeacherSurvey] TS ON TI.tID = TS.tID
    INNER JOIN SchoolSurvey SS ON TS.ssID = SS.ssID
    LEFT OUTER JOIN lkpTeacherQual TQ ON TS.[tchEdQual] = TQ.codeCode
    LEFT OUTER JOIN lkpTeacherQual TQ2 ON TS.[tchQual] = TQ2.codeCode
    WHERE TQ.codeDescription IS NOT NULL
      AND TQ2.codeDescription IS NOT NULL
)

SELECT *
FROM RankedTeachers
WHERE rn = 1;
"""

# Run the query and get the result in a DataFrame
df_teacher_recent_survey_data = pd.read_sql_query(query, engine)
# %store df_teacher_recent_survey_data

# Preview the result
df_teacher_recent_survey_data.head()
