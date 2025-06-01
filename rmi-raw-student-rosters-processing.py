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
import numpy as np
import xlwings as xw
import shutil

def load_config(config_path="config.json"):
    """Load configuration from a JSON file."""
    with open(config_path, 'r') as file:
        config = json.load(file)
    return config['output_directory'], config['source_workbook_filename'], config['empty_census_workbook_filename'], config['clean_census_workbook_filename']
    

# Test loading configuration
output_directory, source_workbook_filename, empty_census_workbook_filename, clean_census_workbook_filename = load_config()
print("Configuration loaded successfully.")

# Import some lookups
# %store -r core_lookups student_lookups censusworkbook_lookups
# %store -r df_enrolments

# %%
# Load workbook

# Combine the directory and filename
workbook_path = os.path.join(output_directory, source_workbook_filename)

# Load the workbook (but not yet any sheet into memory)
xls = pd.ExcelFile(workbook_path)

# See the available sheet names
print("Sheets found:", xls.sheet_names)


# %%
# Get data from raw rosters

# Build the full path
workbook_path = os.path.join(output_directory, source_workbook_filename)

# Load the Excel file
xls = pd.ExcelFile(workbook_path)

# Specify target sheets
target_sheets = ['Newton', 'Larul', 'Maidher', 'Calvin', 'Charlene']

# Prepare a list to collect valid DataFrames
dfs = []

# Track the expected columns
expected_columns = None

# Load each sheet and validate structure
for sheet in target_sheets:
    df = xls.parse(sheet)

    # Initialize expected columns from the first sheet
    if expected_columns is None:
        expected_columns = list(df.columns)
    else:
        # Check if the columns match exactly
        if list(df.columns) != expected_columns:
            raise ValueError(f"Sheet '{sheet}' does not match expected columns: {expected_columns}")

    # Add source sheet name
    df['SourceSheet'] = sheet

    dfs.append(df)

# Concatenate all validated DataFrames
combined_df = pd.concat(dfs, ignore_index=True)

# Display basic preview
display(combined_df.head(3))

# Enhanced summary
total_rows = combined_df.shape[0]
column_list = list(combined_df.columns)

print(f"✅ Successfully combined {len(dfs)} sheets into one DataFrame with {total_rows} total rows.\n")
print(f"🧾 Columns in combined_df ({len(column_list)}):")
print(column_list)

# Summary per SchoolName (if present)
if 'SchoolName' in combined_df.columns:
    print("\n📊 Row counts per SchoolName:")
    print(combined_df['SchoolName'].value_counts(dropna=False).sort_index())
else:
    print("ℹ️ 'SchoolName' column not found in the combined DataFrame.")


# %%
# Identify all rows with NaN and junk
# Identify columns to check (exclude '0')
columns_to_check = ['SchoolName', 'SchoolYear', 'Grade', 'FirstName', 'MiddleInitial', 'LastName', 'Gender', 'Sex', 'BirthDate']

# Create a mask where we treat empty strings and whitespace as NaN
df_check = combined_df[columns_to_check].replace(r'^\s*$', pd.NA, regex=True)

# Flag junk rows: all columns (except '0') are NA or empty
junk_rows_mask = df_check.isna().all(axis=1)

# Summary before dropping
print(f"🗑️ Identified {junk_rows_mask.sum()} junk row(s) where only column '0' has content.")

# Drop them
combined_df_cleaned = combined_df[~junk_rows_mask].copy()

# Confirm cleanup
print(f"✅ Cleaned DataFrame now has {combined_df_cleaned.shape[0]} rows (was {combined_df.shape[0]}).")

# Display rows where SchoolName is still missing
nan_schoolname_rows = combined_df_cleaned[combined_df_cleaned['SchoolName'].isna()]

# Summary and preview
print(f"⚠️ There are {nan_schoolname_rows.shape[0]} row(s) with missing 'SchoolName' still present.")
display(nan_schoolname_rows.head(10))

# Subset of rows where SchoolName is NaN
nan_schoolname_rows = combined_df_cleaned[combined_df_cleaned['SchoolName'].isna()]

# Identify columns that are not entirely NaN in these rows
non_empty_columns = nan_schoolname_rows.dropna(axis=1, how='all').columns

# Display the subset with only those columns
print(f"📌 Showing {len(non_empty_columns)} columns with data in rows with missing 'SchoolName':")
display(nan_schoolname_rows[non_empty_columns].head(10))



# %%
# Clean schNo

# Step 1: Extract valid school codes from lookup
valid_school_codes = {entry['C'] for entry in core_lookups['schoolCodes']}

# Step 2: Flag invalid school codes in the DataFrame
combined_df_cleaned['Invalid_schNo'] = ~combined_df_cleaned['schNo'].isin(valid_school_codes)

# Step 3: Get a list of unique invalid school codes
invalid_schNo_values = combined_df_cleaned.loc[combined_df_cleaned['Invalid_schNo'], 'schNo'].unique().tolist()
total_invalid_rows = combined_df_cleaned['Invalid_schNo'].sum()

# Step 4: Summary
print(f"⚠️ Flagged {total_invalid_rows} row(s) with invalid school codes.")
print(f"⚠️ Found {len(invalid_schNo_values)} unique invalid schNo value(s):")
print(invalid_schNo_values)

# Step 5: Build a lookup dictionary from school name to school code
name_to_code = {entry['N']: entry['C'] for entry in core_lookups['schoolCodes']}

# Step 6: Try to correct schNo based on SchoolName, only for invalid rows
def correct_schno(row):
    if row['Invalid_schNo']:
        school_name = row.get('SchoolName')
        corrected_code = name_to_code.get(school_name)
        if corrected_code:
            return corrected_code
    return row['schNo']  # Leave original if not corrected

# Step 7: Apply correction
combined_df_cleaned['Corrected_schNo'] = combined_df_cleaned.apply(correct_schno, axis=1)

# Step 8: Optional: Re-check how many are still invalid after correction
combined_df_cleaned['Corrected_Invalid_schNo'] = ~combined_df_cleaned['Corrected_schNo'].isin(name_to_code.values())

# Step 9: Summary
fixed_count = combined_df_cleaned['Invalid_schNo'].sum() - combined_df_cleaned['Corrected_Invalid_schNo'].sum()
print(f"✅ Automatically corrected {fixed_count} invalid schNo values based on SchoolName match.")


# %%
# ✅ Clean SchoolName

# Step 1: Extract valid school names from lookup
valid_school_names = {entry['N'].strip() for entry in core_lookups['schoolCodes'] if entry['N']}

# Step 2: Flag invalid school names in the DataFrame
combined_df_cleaned['Invalid_SchoolName'] = ~combined_df_cleaned['SchoolName'].isin(valid_school_names)

# Step 3: Get a list of unique invalid school names
invalid_schoolname_values = (
    combined_df_cleaned.loc[combined_df_cleaned['Invalid_SchoolName'], 'SchoolName']
    .dropna()
    .unique()
    .tolist()
)
total_invalid_names = combined_df_cleaned['Invalid_SchoolName'].sum()

# Step 4: Summary
print(f"⚠️ Flagged {total_invalid_names} row(s) with invalid SchoolName.")
print(f"⚠️ Found {len(invalid_schoolname_values)} unique invalid SchoolName value(s):")
print(invalid_schoolname_values)

# Step 5: Try to correct SchoolName based on fuzzy match
import difflib

def correct_schoolname(name):
    if pd.isna(name):
        return name
    best_match = difflib.get_close_matches(name.strip(), valid_school_names, n=1, cutoff=0.7)
    return best_match[0] if best_match else name  # fallback to original if no good match

# Step 6: Apply correction to only invalid names
combined_df_cleaned['CleanedSchoolName'] = combined_df_cleaned.apply(
    lambda row: correct_schoolname(row['SchoolName']) if row['Invalid_SchoolName'] else row['SchoolName'],
    axis=1
)

# Step 7: Optional: Re-check how many are still invalid
combined_df_cleaned['Corrected_Invalid_SchoolName'] = ~combined_df_cleaned['CleanedSchoolName'].isin(valid_school_names)

# Step 8: Summary
fixed_names_count = combined_df_cleaned['Invalid_SchoolName'].sum() - combined_df_cleaned['Corrected_Invalid_SchoolName'].sum()
print(f"✅ Automatically corrected {fixed_names_count} invalid SchoolName values based on fuzzy match.")


# %%
# Clean SchoolYear

# Step 1: Define expected value
expected_school_year = '2024-2025'

# Step 2: Flag invalid or missing SchoolYear values
invalid_school_year_mask = combined_df_cleaned['SchoolYear'] != expected_school_year

# Step 3: Create a DataFrame with problematic rows
invalid_school_year_df = combined_df_cleaned[invalid_school_year_mask].copy()

# Step 4: Summary
print(f"📅 Total rows in combined_df_cleaned: {combined_df_cleaned.shape[0]}")
print(f"⚠️ Found {invalid_school_year_df.shape[0]} row(s) with invalid or missing SchoolYear.")

# Step 5: Optional preview
if not invalid_school_year_df.empty:
    print("\n📋 Top 10 rows with invalid SchoolYear:")
    display(invalid_school_year_df[['SchoolName', 'SchoolYear', 'SourceSheet']].head(10))
    
    print("\n📊 Count of invalid SchoolYear rows per SourceSheet:")
    print(invalid_school_year_df['SourceSheet'].value_counts().sort_index())

# Step 6: Replace any non-matching or missing values with the correct one
# Define the final cleaned format for SchoolYear
formatted_school_year = f"SY{expected_school_year}"
combined_df_cleaned['SchoolYear'] = combined_df_cleaned['SchoolYear'].where(
    combined_df_cleaned['SchoolYear'] == formatted_school_year,
    formatted_school_year
)

# Confirm result
unique_years = combined_df_cleaned['SchoolYear'].unique()
print(f"✅ All SchoolYear values set to '{expected_school_year}'. Unique values now: {unique_years}")


# %%
# Clean Grade levels
# Step 1: Extract valid grade names from the lookup
valid_grades = {entry['N'] for entry in core_lookups['levels']}

# Step 2: Strip whitespace and unify formatting
combined_df_cleaned['Grade'] = combined_df_cleaned['Grade'].astype(str).str.strip()

# Step 3: Flag invalid grade values
combined_df_cleaned['Invalid_Grade'] = ~combined_df_cleaned['Grade'].isin(valid_grades)

# Step 4: Subset invalid rows
invalid_grades_df = combined_df_cleaned[combined_df_cleaned['Invalid_Grade']].copy()

# Step 5: Summary
print(f"🏫 Total rows: {combined_df_cleaned.shape[0]}")
print(f"⚠️ Found {invalid_grades_df.shape[0]} row(s) with invalid or missing Grade.")

if not invalid_grades_df.empty:
    display(invalid_grades_df[['SchoolName', 'Grade', 'SourceSheet']].head(10))
    print("\n📊 Count of invalid Grade values:")
    print(invalid_grades_df['Grade'].value_counts(dropna=False))

# Step 6: Apply corrections to known mistakes
grade_corrections = {
    'Kiner': 'Kinder',
    'Prek': 'Pre-K',
    'pre-k': 'Pre-K',
    'Grade1': 'Grade 1',
    'grade 1': 'Grade 1',
    '1st grade': 'Grade 1',
    # Add more mappings as needed
}

# Apply correction
combined_df_cleaned['Grade'] = combined_df_cleaned['Grade'].replace(grade_corrections)

# Re-flag invalids after applying corrections
combined_df_cleaned['Invalid_Grade'] = ~combined_df_cleaned['Grade'].isin(valid_grades)

# Re-check how many remain invalid
remaining_invalids = combined_df_cleaned['Invalid_Grade'].sum()
print(f"\n🔄 After applying corrections, {remaining_invalids} row(s) still have invalid Grade values.")


# %%
# Clean and Validate Gender
# Step 1: Define valid values
valid_genders = {'Male', 'Female'}

# Step 2: Strip whitespace and standardize type
combined_df_cleaned['Gender'] = combined_df_cleaned['Gender'].astype(str).str.strip()

# Step 3: Flag invalid values
combined_df_cleaned['Invalid_Gender'] = ~combined_df_cleaned['Gender'].isin(valid_genders)

# Step 4: Show initial invalids
initial_invalids = combined_df_cleaned[combined_df_cleaned['Invalid_Gender']]
print(f"⚠️ Found {initial_invalids.shape[0]} row(s) with invalid Gender before correction.")

# Step 5: Apply basic corrections
gender_corrections = {
    'M': 'Male',
    'F': 'Female',
    'm': 'Male',
    'f': 'Female',
    'male': 'Male',
    'female': 'Female',
    'Femaler': 'Female',
    'Feamale': 'Female',
    'MALE': 'Male',
    'FEMALE': 'Female',
    'John': 'Male',
}

combined_df_cleaned['Gender'] = combined_df_cleaned['Gender'].replace(gender_corrections)

# Step 6: Re-flag invalids after correction
combined_df_cleaned['Invalid_Gender'] = ~combined_df_cleaned['Gender'].isin(valid_genders)

# Step 7: Final summary
remaining_invalids = combined_df_cleaned['Invalid_Gender'].sum()
print(f"✅ Gender correction complete. Remaining invalid rows: {remaining_invalids}")

# Optional: Preview a few of the remaining invalids
if remaining_invalids:
    display(combined_df_cleaned[combined_df_cleaned['Invalid_Gender']][['SchoolName', 'Gender', 'SourceSheet']].head(10))


# %%
# BirthDate Cleaning and Age Flagging (Using Existing Context)

# Step 1: Extract census year info for 2025
census_info = next(x for x in censusworkbook_lookups['censusYears'] if x['svyYear'] == 2025)
census_date = pd.to_datetime(census_info['svyCensusDate'])
ps_age = census_info['svyPSAge']  # Age for Grade 1

# Step 2: Build expected age mapping from grade names
grade_age_map = {
    entry['N']: ps_age-1 + entry['YoEd']
    for entry in core_lookups['levels']
    if 'YoEd' in entry and pd.notna(entry['YoEd'])
}

# Summary: show census date, ps_age, and sample grade mappings
print(f"📅 Census Date: {census_date.date()}")
print(f"🎯 Official Age for Grade 1: {ps_age}")
print(f"📘 Grade-to-Expected-Age Mapping (first 5):")
for grade, age in list(grade_age_map.items())[:5]:
    print(f"   - {grade}: {age} years old")

# Step 3: Parse BirthDate to datetime safely
combined_df_cleaned['SourceParsedBirthDate'] = pd.to_datetime(
    combined_df_cleaned['BirthDate'], errors='coerce'
)

# Step 4: Calculate actual age at census date
combined_df_cleaned['SourceAgeAtCensus'] = combined_df_cleaned['SourceParsedBirthDate'].apply(
    lambda dob: census_date.year - dob.year - ((census_date.month, census_date.day) < (dob.month, dob.day))
    if pd.notna(dob) else None
)

# Step 5: Determine expected age based on Grade
combined_df_cleaned['ExpectedAge'] = combined_df_cleaned['Grade'].map(grade_age_map)

# Step 6: Flag rows with unrealistic age difference (±3+ years from expected)
combined_df_cleaned['AgeFlagged'] = combined_df_cleaned.apply(
    lambda row: (
        pd.notna(row['ExpectedAge']) and
        pd.notna(row['SourceAgeAtCensus']) and
        abs(row['SourceAgeAtCensus'] - row['ExpectedAge']) >= 3
    ),
    axis=1
)

# Step 7: Summary
flagged_rows = combined_df_cleaned[combined_df_cleaned['AgeFlagged']]
print(f"⚠️ Found {flagged_rows.shape[0]} student(s) with BirthDate far from expected for their Grade.")

# Optional preview
display(flagged_rows[['SchoolName', 'FirstName', 'LastName', 'Grade', 'BirthDate', 'SourceParsedBirthDate', 'SourceAgeAtCensus', 'ExpectedAge', 'SourceSheet']].head(10))

# Step 8: Compare reported Age column with calculated AgeAtCensus
# Make sure both are numeric and not null
combined_df_cleaned['SourceReportedAge'] = pd.to_numeric(combined_df_cleaned['Age'], errors='coerce')

combined_df_cleaned['AgeMismatch'] = combined_df_cleaned.apply(
    lambda row: (
        pd.notna(row['SourceReportedAge']) and
        pd.notna(row['SourceAgeAtCensus']) and
        row['SourceReportedAge'] != row['SourceAgeAtCensus']
    ),
    axis=1
)

# Summary of mismatches
mismatched_age_df = combined_df_cleaned[combined_df_cleaned['AgeMismatch']]
print(f"🔍 Found {mismatched_age_df.shape[0]} row(s) where reported Age (calculated in the RMI raw rosters) differs from calculated AgeAtCensus.")

# Optional preview of mismatches
display(mismatched_age_df[['SchoolName', 'FirstName', 'LastName', 'Grade', 'BirthDate', 'SourceParsedBirthDate', 'SourceAgeAtCensus', 'ExpectedAge', 'SourceReportedAge', 'SourceSheet']].head(10))

# Step 9 (revised): Fix unrealistic or missing birthdates into FixedParsedBirthDate
def get_default_birthdate(row):
    if pd.isna(row['ExpectedAge']):
        return None
    expected_birth_year = census_date.year - row['ExpectedAge']
    return pd.Timestamp(f"{expected_birth_year}-01-01")

# Create the new column (default to original value)
combined_df_cleaned['FixedParsedBirthDate'] = combined_df_cleaned['SourceParsedBirthDate']

# Identify rows to fix
birthdate_fix_mask = combined_df_cleaned['AgeFlagged'] | combined_df_cleaned['SourceParsedBirthDate'].isna()

# Apply fix only to those rows
combined_df_cleaned.loc[birthdate_fix_mask, 'FixedParsedBirthDate'] = combined_df_cleaned.loc[
    birthdate_fix_mask
].apply(get_default_birthdate, axis=1)

# Add 'DoB Estimate' flag where we applied a fix
combined_df_cleaned['DoB Estimate'] = ''
combined_df_cleaned.loc[birthdate_fix_mask, 'DoB Estimate'] = 'Yes'

# Recalculate AgeAtCensus using the fixed birthdates, and store in a new column
combined_df_cleaned['FixedAgeAtCensus'] = combined_df_cleaned['FixedParsedBirthDate'].apply(
    lambda dob: census_date.year - dob.year - ((census_date.month, census_date.day) < (dob.month, dob.day))
    if pd.notna(dob) else None
)

# Summary
print(f"🛠️ Fixed {birthdate_fix_mask.sum()} BirthDate value(s) using January 1 of expected birth year (based on grade).")
print(f"ℹ️ AgeFlagged column retained to indicate which rows were originally flagged.")


# %%
import matplotlib.pyplot as plt

# Helper function to categorize age groups based on provided age column
def categorize_age_group(df, age_column):
    return df.apply(
        lambda row: (
            'Official Age' if pd.notna(row[age_column]) and pd.notna(row['ExpectedAge']) and row[age_column] == row['ExpectedAge']
            else 'Under Age' if pd.notna(row[age_column]) and pd.notna(row['ExpectedAge']) and row[age_column] < row['ExpectedAge']
            else 'Over Age' if pd.notna(row[age_column]) and pd.notna(row['ExpectedAge']) and row[age_column] > row['ExpectedAge']
            else 'Unknown'
        ),
        axis=1
    )

# --- Plot using SourceAgeAtCensus ---
combined_df_cleaned['SourceAgeGroup'] = categorize_age_group(combined_df_cleaned, 'SourceAgeAtCensus')

grade_group_source = combined_df_cleaned[
    combined_df_cleaned['Grade'].notna() & combined_df_cleaned['SourceAgeGroup'].notna()
].groupby(['Grade', 'SourceAgeGroup']).size().unstack(fill_value=0)

grade_order = [entry['N'] for entry in core_lookups['levels'] if entry['N'] in grade_group_source.index]
grade_group_source = grade_group_source.reindex(grade_order)

# Reorder age group columns
age_order = ['Under Age', 'Official Age', 'Over Age']
grade_group_source = grade_group_source[[col for col in age_order if col in grade_group_source.columns]]

# Plot 1: All age groups (Source)
ax1 = grade_group_source.plot(
    kind='bar', stacked=False, figsize=(12, 6),
    title='(Before Cleanup) Enrollment by Grade and Age Group',
    ylabel='Number of Students', xlabel='Grade',
    rot=45, grid=True, legend=True
)
for container in ax1.containers:
    ax1.bar_label(container, label_type='edge', fontsize=8)
plt.tight_layout()
plt.show()

# Plot 2: Only Under and Over Age (Source)
source_filtered = grade_group_source.drop(columns='Official Age', errors='ignore')
ax2 = source_filtered.plot(
    kind='bar', stacked=False, figsize=(12, 6),
    title='(Before Cleanup) Enrollment by Grade: Under & Over Age Only',
    ylabel='Number of Students', xlabel='Grade',
    rot=45, grid=True, legend=True
)
for container in ax2.containers:
    ax2.bar_label(container, label_type='edge', fontsize=8)
plt.tight_layout()
plt.show()

# --- Plot using FixedAgeAtCensus ---
combined_df_cleaned['AgeGroup'] = categorize_age_group(combined_df_cleaned, 'FixedAgeAtCensus')

grade_group_fixed = combined_df_cleaned[
    combined_df_cleaned['Grade'].notna() & combined_df_cleaned['AgeGroup'].notna()
].groupby(['Grade', 'AgeGroup']).size().unstack(fill_value=0)

grade_group_fixed = grade_group_fixed.reindex(grade_order)
grade_group_fixed = grade_group_fixed[[col for col in age_order if col in grade_group_fixed.columns]]

# Plot 3: All age groups (Fixed)
ax3 = grade_group_fixed.plot(
    kind='bar', stacked=False, figsize=(12, 6),
    title='(After Cleanup) Enrollment by Grade and Age Group',
    ylabel='Number of Students', xlabel='Grade',
    rot=45, grid=True, legend=True
)
for container in ax3.containers:
    ax3.bar_label(container, label_type='edge', fontsize=8)
plt.tight_layout()
plt.show()

# Plot 4: Only Under and Over Age (Fixed)
fixed_filtered = grade_group_fixed.drop(columns='Official Age', errors='ignore')
ax4 = fixed_filtered.plot(
    kind='bar', stacked=False, figsize=(12, 6),
    title='(After Cleanup) Enrollment by Grade: Under & Over Age Only',
    ylabel='Number of Students', xlabel='Grade',
    rot=45, grid=True, legend=True
)
for container in ax4.containers:
    ax4.bar_label(container, label_type='edge', fontsize=8)
plt.tight_layout()
plt.show()


# %%
display(combined_df_cleaned.iloc[:, 11:20])  # columns 1–5

# %%
# Work on the From column

# Step 1: Define initial mapping (you can expand this later)
from_mapping = {
    'Transferred IN': 'Transferred In',
    'Transfer': 'Transferred In',
    'Transfer From': 'Transferred In',
    'Transfer In': 'Transferred In',
    'Transfer in': 'Transferred In',
    'Transferred In': 'Transferred In',
    'Transferred in': 'Transferred In',
    'Transferred IN': 'Transferred In',
    '0': 'Transferred In',
    'Carlos': 'Transferred In',
    'EKN': 'Transferred In',
    'EPES': 'Transferred In',
    'New': 'New Enrolment',
    'new': 'New Enrolment',
    'NEW': 'New Enrolment',    
    'Repeat': 'Repeater',
    'Repeater': 'Repeater',
    'Retained': 'Repeater',
    'Retain': 'Repeater',
    'Retained': 'Repeater',
    'Retuening': np.nan,
    'Returing': np.nan,
    'Return/Retained': np.nan,
    'Returning': np.nan,
    'returning': np.nan,
    'Returning': np.nan,
    'Reurning': np.nan,
    'Re-Entry': 'Repeater',
    'Returned': np.nan,
    'Continuing': np.nan,
    'Continung': np.nan,
    'Contiuing': np.nan,
    '': np.nan,
    # Add more mappings as needed...
}

# Step 2: Standardize 'From' values using the mapping
combined_df_cleaned['CleanedFrom'] = combined_df_cleaned['Return/Retained'].astype(str).map(from_mapping)

# Step 3: Show unique raw values in 'Return/Retained' for further mapping
unique_values = combined_df_cleaned['Return/Retained'].dropna().astype(str).unique()

print("📋 Unique values in 'Return/Retained':")
for val in sorted(unique_values, key=lambda x: x.lower()):
    print(f" - '{val}'")

# Step 4: Define final accepted values
valid_from_values = {'ECE', 'New Enrolment', 'Repeater', 'Transferred In'}

# Step 5: Count how many rows have valid 'From' values
from_counts_before = combined_df_cleaned['CleanedFrom'].value_counts(dropna=False)
valid_rows_before = combined_df_cleaned['CleanedFrom'].isin(valid_from_values).sum()
invalid_rows_before = combined_df_cleaned.shape[0] - valid_rows_before

# Step 6: Infer those that came from ECE
# Set ECE values based on an existing enrolment record of previous year.
# In other words, if a matching student df_enrolments -> combined_df_cleaned
# ('stuGiven' = 'FirstName' and 'stuFamilyName' = 'FamilyName' and
# 'stuDoB' = 'FixedParsedBirthDate') has a record in ECE last year
# ('stueClass' = 'GK' and 'stueYear' = '2023-2024'), then set CleanedFrom = 'ECE'

# --- Valid vs Invalid before ECE inference ---
valid_from_values = {'ECE', 'New Enrolment', 'Repeater', 'Transferred In'}
# Count valid values including NaNs
valid_mask = combined_df_cleaned['CleanedFrom'].isin(valid_from_values) | combined_df_cleaned['CleanedFrom'].isna()
valid_rows = valid_mask.sum()
invalid_rows = combined_df_cleaned.shape[0] - valid_rows


print(f"📊 Initial summary of 'CleanedFrom' column cleanup (before ECE inference):")
print(f"✅ Valid 'CleanedFrom' values (ECE, New Enrolment, Repeater, Transferred In): {valid_rows}")
print(f"❌ Invalid or unmapped 'CleanedFrom' values: {invalid_rows}")
print(f"📋 Frequency breakdown:")
print(from_counts)

# --- Inference from prior year ECE enrollment ---
original_ece_count = (combined_df_cleaned['CleanedFrom'] == 'ECE').sum()

# Safely calculate previous year
start_year = int(expected_school_year.split('-')[1])
previous_year = start_year - 1

ece_last_year = df_enrolments[
    (df_enrolments['stueClass'] == 'GK') &
    (df_enrolments['stueYear'] == previous_year)
]

ece_lookup = set(
    zip(
        ece_last_year['stuGiven'].str.upper().str.strip(),
        ece_last_year['stuFamilyName'].str.upper().str.strip(),
        pd.to_datetime(ece_last_year['stuDoB'], errors='coerce')
    )
)

def infer_from_ece(row):
    if pd.notna(row['CleanedFrom']):
        return row['CleanedFrom']
    key = (str(row['FirstName']).upper().strip(), str(row['LastName']).upper().strip(), pd.to_datetime(row['FixedParsedBirthDate'], errors='coerce'))
    if key in ece_lookup:
        return 'ECE'
    return row['CleanedFrom']

combined_df_cleaned['CleanedFrom'] = combined_df_cleaned.apply(infer_from_ece, axis=1)

# --- Final ECE summary ---
ece_total = (combined_df_cleaned['CleanedFrom'] == 'ECE').sum()
ece_inferred = ece_total - original_ece_count

print(f"\n🧠 ECE inference summary:")
print(f"🟢 Already had ECE before inference: {original_ece_count}")
print(f"🔍 Inferred ECE entries: {ece_inferred}")
print(f"✅ Total ECE entries after processing: {ece_total}")


# %%
# ✅ Infer Attended ECE Status (anytime in the past or from source column)
# Match based on FirstName, LastName, and FixedParsedBirthDate
# If student appears in df_enrolments with stueClass == 'GK', OR has 'Yes' in original ECE/Kinder field, mark as 'Yes'

# Step 1: Prepare matching key
combined_df_cleaned['__MatchKey'] = (
    combined_df_cleaned['FirstName'].astype(str).str.strip().str.upper() + '|' +
    combined_df_cleaned['LastName'].astype(str).str.strip().str.upper() + '|' +
    combined_df_cleaned['FixedParsedBirthDate'].astype(str)
)

df_enrolments['__MatchKey'] = (
    df_enrolments['stuGiven'].astype(str).str.strip().str.upper() + '|' +
    df_enrolments['stuFamilyName'].astype(str).str.strip().str.upper() + '|' +
    df_enrolments['stuDoB'].astype(str)
)

# Step 2: Match keys from ECE history (GK class)
ece_match_keys = set(
    df_enrolments.loc[df_enrolments['stueClass'] == 'GK', '__MatchKey']
)

# Step 3: Create inferred column (start with None)
combined_df_cleaned['InferredAttendedECE'] = None

# Step 4: Source-based assignment
ece_col = 'Attended \nECE/Kinder\n ?'
source_yes_mask = pd.Series([False] * len(combined_df_cleaned))
if ece_col in combined_df_cleaned.columns:
    source_yes_mask = combined_df_cleaned[ece_col].astype(str).str.strip().str.upper().isin({'YES', 'Y'})
    combined_df_cleaned.loc[source_yes_mask, 'InferredAttendedECE'] = 'Yes'

# Step 5: Inference from enrollment history (only if not already 'Yes')
inferred_yes_mask = (
    combined_df_cleaned['__MatchKey'].isin(ece_match_keys) &
    (combined_df_cleaned['InferredAttendedECE'] != 'Yes')
)
combined_df_cleaned.loc[inferred_yes_mask, 'InferredAttendedECE'] = 'Yes'

# Step 6: Cleanup temporary key
combined_df_cleaned.drop(columns='__MatchKey', inplace=True)
df_enrolments.drop(columns='__MatchKey', inplace=True, errors='ignore')

# Step 7: Summary
total_students = len(combined_df_cleaned)
source_yes = source_yes_mask.sum()
inferred_yes = inferred_yes_mask.sum()
final_yes = (combined_df_cleaned['InferredAttendedECE'] == 'Yes').sum()

print(f"🧠 Inferred or recognized ECE attendance (any year):")
print(f"📌 'Yes' from source column '{ece_col}': {source_yes}")
print(f"🔍 Inferred from historical GK enrollment: {inferred_yes}")
print(f"✅ Total students marked as having attended ECE: {final_yes} of {total_students}")


# %%
# ✅ Clean and Infer Transferred From which school
# Step 1: Define draft mapping (to be refined after inspecting unique values)
transferred_from_mapping = {
    '#REF!': 'None',
    'A.E.S': 'Aur Elementary School',
    'AES': 'Aur Elementary School',
    'AHS': 'Assumption High School',
    'Aelok Elementary School': 'Aerok Elementary School',
    'Aelonlaplap Elementary Schools': 'Delap Elementary School',
    'Aelonlaplap, Buoj Elementary School': 'Buoj Elementary School',
    'Aelonlaplap, Buoj Elementary School.': 'Buoj Elementary School',
    'Aelonlaplap, Je ES': 'Jobwan Elementary School',
    'Aelonlaplap, Jebwon Elementary School': 'Jobwan Elementary School',
    'Aelonlaplap, Katiej Elementary School': 'Katiej Elementary School',
    'Aelonlaplap, Mejel Elementary School': 'Mejel Elementary School',
    'Aelonlaplap, Woja ES.': 'Woja Elementary School (Ailinglaplap)',
    'Aiea HS( Hawaii': 'International',
    'Ailuk E.S.': 'Ailuk Elementary School',
    'Ailuk Elementary School': 'Ailuk Elementary School',
    'Airok E.S.(MAL)': 'Aerok Elementary School',
    'Airok Elem. School(Ailinglaplap)': 'Woja Elementary School (Ailinglaplap)',
    'Ajeltake': 'Ajeltake Elementary School',
    'Ajeltake E.S.': 'Ajeltake Elementary School',
    'Ajeltake Elem.': 'Ajeltake Elementary School',
    'Ajeltake Elementary School': 'Ajeltake Elementary School',
    'Alaska': 'International',
    'Arkansas': 'International',
    'Arno': 'Arno Elementary School',
    'Arno E.S.': 'Arno Elementary School',
    'Arno Elem. School': 'Arno Elementary School',
    'Arno Elementary School': 'Arno Elementary School',
    'Arno,  Ine ES.': 'Ine Elementary School',
    'Arno, Ine ES.': 'Ine Elementary School',
    'Assumption Elementary School': 'Assumption Elementary School',
    'Assumption H.S.': 'Assumption High School',
    'Aur Elem. School': 'Aur Elementary School',
    'Aur Elementary School': 'Aur Elementary School',
    'Aur, Tobal Elementary School': 'Tobal Elementary School',
    'Bikareej Elementary School': 'Bikarej Elementary School',
    'Bikarej Elem. School': 'Bikarej Elementary School',
    'Buoj E.S.': 'Buoj Elementary School',
    'Buoj Elem. School': 'Buoj Elementary School',
    'Buoj Elementary School': 'Buoj Elementary School',
    'COHS': 'Majuro Coop High School',
    'COOP': 'Majuro Coop High School',
    'COOP HS': 'Majuro Coop High School',
    'Calvary High School': 'Ebeye Calvary High School',
    'Carfield Elem. USA': 'International',
    'Carter HS': 'International',
    'Carter HS (USA)': 'International',
    'Cascade Middle SChool': 'International',
    'Cascade Middle School': 'International',
    'Castle High, HI': 'International',
    'Coop': 'Majuro Coop High School',
    'Coop HS': 'Majuro Coop High School',
    'DES': 'Delap Elementary School',
    'Deaf Center': 'Ebeye Deaf Center Primary',
    'Delap Elem. School': 'Delap Elementary School',
    'Delap Elementary School': 'Delap Elementary School',
    'Dropped-out': 'None',
    'ECES': 'Ebeye Christian Elementary School',
    'EPES': 'Ebeye Public Elementary School',
    'EPMS': 'Ebeye Public Middle School',
    'EPSS, Ebeye': 'Ebeye Public Elementary School',
    'Ebeye': 'Ebeye Public Elementary School',
    'Ebeye Calvary Elementary School': 'Ebeye Calvary Elementary School',
    'Ebeye Calvary H.S.': 'Ebeye Calvary High School',
    'Ebeye E. School': 'Ebeye Public Elementary School',
    'Ebeye Elem. School': 'Ebeye Public Elementary School',
    'Ebeye Elem.School': 'Ebeye Public Elementary School',
    'Ebeye Elementary School': 'Ebeye Public Elementary School',
    'Ebeye Middle Scool': 'Ebeye Public Middle School',
    'Ebeye Public E. School': 'Ebeye Public Middle School',
    'Ebeye Public E.S.': 'Ebeye Public Middle School',
    'Ebeye Public Elementary School': 'Ebeye Public Elementary School',
    'Ebeye SDA': 'Ebeye SDA Elementary School',
    'Ebeye, ECES': 'Ebeye Christian Elementary School',
    'Ebeye, PSS': 'Ebeye SDA High School',
    'Ebeye, SDA': 'Ebeye SDA Elementary School',
    'Ebezon Elem PI': 'International',
    'Ebon E.S.': 'Ebeye Public Elementary School',
    'Ebon Elem. School': 'Ebon Elementary School',
    'Ebon, Eneko Ion Elementary School': 'Enekoion Elementary School',
    'Ejit Elem. School': 'Ejit Elementary School',
    'Ejit Elementary School': 'Ejit Elementary School',
    'Ejit Is. Elem School': 'Ejit Elementary School',
    'Enejet Elementary School': 'Enejet Elementary School',
    'Enejet Elemntary School': 'Enejet Elementary School',
    'Enekion Elem Ebon': 'Enekoion Elementary School',
    'Enewetak ES': 'Enewetak Elementary School',
    'Enewetak Elem.': 'Enewetak Elementary School',
    'Enewetak Elementary School': 'Enewetak Elementary School',
    'FIFe HS (WA State)': 'International',
    'Father Hacker H.S.': 'Father Hacker High School',
    'Fife HS, WA State': 'International',
    'Gateway Elem. AZ': 'International',
    'Georgia Middle USA': 'International',
    'HI': 'International',
    'Hawaii': 'International',
    'Hillside Elem. WA': 'International',
    'Hilo High School': 'Jabro High School',
    'Hilo Middle  Hilo': 'International',
    'Ine Elem.': 'Ine Elementary School',
    'Ine Elementary SChool': 'Ine Elementary School',
    'Ine Elementary School': 'Ine Elementary School',
    'JHS': 'Jaluit High School',
    'JPS': 'None',
    'Jabor E.S.': 'Jabor Elementary School',
    'Jabor Elementary School': 'Jabor Elementary School',
    'Jabro HS (EBEYE)': 'Jabro High School',
    'Jabro Private School': 'Jabro High School',
    'Jabro School': 'Jabro High School',
    'Jaluit Elem. School': 'Jaluit Elementary School',
    'Jaluit High School': 'Jaluit High School',
    'Jaluit Saint Joseph Elementary School': 'St. Joseph Elementary School',
    'Jaluit, Mejrirok Elementary School': 'Mejrirok Elementary School',
    'Jang Elementary School': 'Jang Elementary School',
    'Japo Arno Elem.': 'Japo Elementary School',
    'Japo Elementary School': 'Japo Elementary School',
    'Jeh Elem. School': 'Jeh Elementary School',
    'Jeh Elementary School': 'Jeh Elementary School',
    'Jepo Elem. School': 'Japo Elementary School',
    'Jobwan Elemntary School': 'Jobwan Elementary School',
    'KAHS': 'Kwajalein Atoll High School',
    'Kalenia Ole Elem. HI': 'International',
    'Kamaikin High, WA': 'International',
    'Kaven Elementary School': 'Kaven Elementary School',
    'Kilange Elementary School': 'Kilange Elementary School',
    'Kili E.S.': 'Kili Elementary School',
    'Kili Elem.': 'Kili Elementary School',
    'Kili Elementary School': 'Kili Elementary School',
    'Kili Island Elementary School': 'Kili Elementary School',
    'Kwajalein Atoll H.S.': 'Kwajalein Atoll High School',
    'Kwajalein Atoll High School': 'Kwajalein Atoll High School',
    'LES': 'Lukonwod Elementary School',
    'LHS': 'Laura High School',
    'LIES': 'Long Island Elementary School',
    'LSA': 'Life Skills Academy',
    'Lae Elementary School': 'Lae Elementary School',
    'Laura Elementary School': 'Laura Elementary School',
    'Laura H.S.': 'Laura High School',
    'Light House': 'Lighthouse Apostolic Academy',
    'Likiep Elem. School': 'Likiep Elementary School',
    'Long Isl. Elementary School': 'Long Island Elementary School',
    'Long Island E.S.': 'Long Island Elementary School',
    'Long Island Elementary School': 'Long Island Elementary School',
    'Longar Elementary School': 'Longar Elementary School',
    'Lonng Island Elem. School': 'Long Island Elementary School',
    'Lukoj Elem. School': 'Lukoj Elementary School',
    'Lukwonwod Elem.': 'Lukonwod Elementary School',
    'MBCA': 'International',
    'MCHS': 'Marshall Christians High School',
    'MIHS': 'Marshall Islands High School',
    'MMS': 'Majuro Middle School',
    'Mae Elementary SChool': 'Mae Elementary School',
    'Majkin Elementary school': 'Majkin Elementary School',
    'Majuro': 'None',
    'Majuro Baptist Christian Academy': 'Laura Christian Academy',
    'Majuro Middle School': 'Majuro Middle School',
    'Majuro Midle School': 'Majuro Middle School',
    'Majuro SDA': 'Delap SDA Elementary School',
    'Majuro, Ajeltake Elementary School': 'Ajeltake Elementary School',
    'Majuro, Assumption Primary School': 'None',
    'Majuro, Delap Elementary School': 'Majuro Coop Elementary School',
    'Majuro, Ejit Elementary School': 'Majuro Baptist Elementary School',
    'Majuro, Long Island Elementary School': 'Long Island Elementary School',
    'Majuro, MBCA Primary School': 'None',
    'Majuro, North Delap Elementary School': 'North Delap Elementary School',
    'Majuro, Rairok Rainbow Elementary School': 'Majuro Coop Elementary School',
    'Majuro, SDA Primary School': 'None',
    'Makapala Elem. HI': 'International',
    'Maloelap': 'Aerok Elementary School',
    'Maloelap Elementary Schools': 'Aerok Elementary School',
    'Marshall Christian Elem. School': 'Marshall Christians High School',
    'Marshall Christian High School': 'Marshall Christians High School',
    'Marshall Islands H.S': 'Marshall Islands High School',
    'Marshall Islands High School': 'Marshall Islands High School',
    'Matoleen Elementary School': 'Matolen Elementary School',
    'Matolen Elementary School': 'Matolen Elementary School',
    'Maui HS (Hawaii)': 'International',
    'Maui High, HI': 'International',
    'McKinley, HI': 'International',
    'Mckay HS': 'International',
    'Mejatto E.S.': 'Mejatto Elementary School',
    'Mejatto Elem.': 'Mejatto Elementary School',
    'Mejatto Elem.School': 'Mejatto Elementary School',
    'Mejit Elem. School': 'Mejit Elementary School',
    'Mejit Elementary School': 'Mejit Elementary School',
    'Mejjato Elementary School': 'Mejatto Elementary School',
    'Mejrirok E.S.': 'Mejrirok Elementary School',
    'Mejrirok Elementary SChool': 'Mejrirok Elementary School',
    'Mejrirok Elementary School': 'Mejrirok Elementary School',
    'Mili Elem.': 'Mili Elementary School',
    'Mili Elem. School': 'Mili Elementary School',
    'Mili Elem.School': 'Mili Elementary School',
    'Mili Elementary School': 'Mili Elementary School',
    'Mili elementary School': 'Mili Elementary School',
    'Mili, Tokewa Elem.': 'Tokewa Elementary School',
    'N Middle School WA': 'International',
    'NDES': 'North Delap Elementary School',
    'NIHS': 'Northern Islands High School',
    'Nallo Elem. School': 'Nallo Elementary School',
    'Nallo Elementary School': 'Nallo Elementary School',
    'Nallu Elem.School': 'Nallo Elementary School',
    'Namdrik': 'Namdrik Elementary School',
    'Namdrik Elementary School': 'Namdrik Elementary School',
    'Namo': 'Namu Elementary School',
    'Namu Elem.School': 'Namu Elementary School',
    'Namu Elementary School': 'Namu Elementary School',
    'Narmij Elem Jaluit': 'Narmij Elementary School',
    'Narmij Elementary School': 'Narmij Elementary School',
    'Niu Valley Elem. HI': 'International',
    'North Carolina USA': 'International',
    'North Delap Elem. School': 'North Delap Elementary School',
    'North Delap Elem.School': 'North Delap Elementary School',
    'North Delap Elementary School': 'North Delap Elementary School',
    'Ollet Elementary School': 'Ollet Elementary School',
    'PPS - Ailinglaplap': 'None',
    'PSS -jaluit': 'Jaluit Elementary School',
    'Pahoa HS. HILO': 'International',
    'Phillippines': 'International',
    'QPS': 'Queen of Peace Elementary School',
    'RES': 'Rita Elementary School',
    'RRES': 'RonRon Protestant Elementary School',
    'Rairok E.S.': 'Rairok Elementary School',
    'Rairok Elem. School': 'Rairok Elementary School',
    'Rairok Elementary School': 'Rairok Elementary School',
    'Rita E.S.': 'Rita Elementary School',
    'Rita Elem. School': 'Rita Elementary School',
    'Rita Elementary School': 'Rita Elementary School',
    'Rongrong E.S.': 'RonRon Protestant Elementary School',
    'SDA': 'None',
    'SDA HS': 'None',
    'Santo PSS': 'None',
    'St.Joseph Academy': 'St. Joseph Elementary School',
    'TPS': 'None',
    'Tarawa Elem.School': 'Tarawa Elementary School',
    'Texas Hs. USA': 'International',
    'Tinak Arno Elem.': 'Tinak Elementary School',
    'Tinak Elem. School': 'Tinak Elementary School',
    'Tinak Elementary School': 'Tinak Elementary School',
    'Tobal Aur Elem.': 'Tobal Elementary School',
    'Toka Elementary School': 'Toka Elementary School',
    'Toka elementary School': 'Toka Elementary School',
    'Tokewa Elementary School': 'Tokewa Elementary School',
    'Tutu Elementary School': 'Tutu Elementary School',
    'U,S': 'International',
    'U.S': 'International',
    'U.S Logan Middle School': 'International',
    'U.S.A': 'International',
    'USA': 'International',
    'Ujae Elem. School': 'Ujae Elementary School',
    'Ujae Elementary School': 'Ujae Elementary School',
    'Ulien Elem. School': 'Ulien Elementary School',
    'Ulien Elementary School': 'Ulien Elementary School',
    'Utrok Elem.School': 'Utrik Elementary School',
    'WES, WA': 'International',
    'WPES': 'None',
    'WPES, Wotje': 'None',
    'WSHS': 'None',
    'Washington, USA': 'International',
    'Webling Elem. HI': 'International',
    'Wichita Elem KS': 'International',
    'Woja Ailinlaplap Elem.': 'Woja Elementary School (Ailinglaplap',
    'Woja Elem.': 'Woja Elementary School (Majuro)',
    'Woja Elem. School': 'Woja Elementary School (Majuro)',
    'Woja Elem.School (Majuro)': 'Woja Elementary School (Majuro)',
    'Wotho Elem.School': 'Wotho Elementary School',
    'Wotje E.S.': 'Wotje Elementary School',
    'Wotje Elementary School': 'Wotje Elementary School',
    'Wotje, Wotje Elementary School': 'Wotje Elementary School',
    'Xavier': 'International',
    'YSP': 'International',
    'from Catholic': 'International',
    'from Hawaii': 'International',
    'from Lae': 'Lae Elementary School',
    'late registered': 'None',
    '': np.nan,
    # Add more mappings after reviewing the unique values below...
}

# Step 2: Apply the mapping to source column
combined_df_cleaned['InferredTransferredFromWhichSchool'] = (
    combined_df_cleaned['Transferred\nFROM']
    .astype(str)
    .str.strip()
    .replace('', np.nan)
    .map(transferred_from_mapping)
)

# Step 3: Show unique raw values for mapping refinement
# print("📋 Unique values in 'Transferred\\nFROM':")
# unique_transferred_from_values = (
#     combined_df_cleaned['Transferred\nFROM']
#     .dropna()
#     .astype(str)
#     .str.strip()
#     .unique()
# )
# for val in sorted(unique_transferred_from_values, key=lambda x: x.lower()):
#     print(f" - '{val}'")

# Step 4: Infer from previous year's enrolment if not already set
start_year = int(expected_school_year.split('-')[1])
previous_year = start_year - 1

# Create student match key
combined_df_cleaned['__MatchKey'] = (
    combined_df_cleaned['FirstName'].astype(str).str.strip().str.upper() + '|' +
    combined_df_cleaned['LastName'].astype(str).str.strip().str.upper() + '|' +
    combined_df_cleaned['FixedParsedBirthDate'].astype(str)
)
df_enrolments['__MatchKey'] = (
    df_enrolments['stuGiven'].astype(str).str.strip().str.upper() + '|' +
    df_enrolments['stuFamilyName'].astype(str).str.strip().str.upper() + '|' +
    df_enrolments['stuDoB'].astype(str)
)

# Create a lookup from last year’s enrolments with their school number
df_last_year = df_enrolments[df_enrolments['stueYear'] == previous_year][['__MatchKey', 'schNo']].drop_duplicates()
last_year_school_lookup = dict(zip(df_last_year['__MatchKey'], df_last_year['schNo']))

# Infer where not already set
mask_missing = combined_df_cleaned['InferredTransferredFromWhichSchool'].isna()
inferred_values = combined_df_cleaned.loc[mask_missing, '__MatchKey'].map(last_year_school_lookup)

# Only keep inference if it's from a *different* school than current
inferred_diff_school = inferred_values[
    inferred_values != combined_df_cleaned.loc[mask_missing, 'Corrected_schNo']
]

combined_df_cleaned.loc[inferred_diff_school.index, 'InferredTransferredFromWhichSchool'] = inferred_diff_school

# Step 5: Cleanup
combined_df_cleaned.drop(columns='__MatchKey', inplace=True)
df_enrolments.drop(columns='__MatchKey', inplace=True, errors='ignore')

# Step 6: Summary
total_cleaned = combined_df_cleaned['InferredTransferredFromWhichSchool'].notna().sum()
total_rows = combined_df_cleaned.shape[0]
total_inferred = inferred_diff_school.notna().sum()
total_cleaned_only = total_cleaned - total_inferred

print("📊 Summary of 'InferredTransferredFromWhichSchool' generation:")
print(f"🧼 Cleaned from source: {total_cleaned_only}")
print(f"🧠 Inferred from previous enrolments: {total_inferred}")
print(f"✅ Total populated: {total_cleaned} out of {total_rows}")


# %%
import difflib

# Step 1: Get unique raw values
raw_transfer_values = (
    combined_df_cleaned['Transferred\nFROM']
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

# Step 2: Reference: official school names and acronyms
school_codes_df = pd.DataFrame(core_lookups['schoolCodes'])  # Convert list of dicts to DataFrame
school_codes_df['N'] = school_codes_df['N'].astype(str).str.strip()
school_codes_df['C'] = school_codes_df['C'].astype(str).str.strip()

official_school_names = school_codes_df['N'].unique().tolist()

# Build acronym mapping: clean acronyms for easy lookup (e.g., LHS → Laura High School)
known_acronym_mapping = {}
for _, row in school_codes_df.iterrows():
    name = row['N']
    acronym = ''.join([word[0] for word in name.split() if word[0].isalpha()]).upper()
    if len(acronym) >= 2:  # Keep only meaningful acronyms
        known_acronym_mapping[acronym] = name

# Step 3: Enhanced draft mapping logic
first_draft_mapping = {}

for raw_value in raw_transfer_values:
    val = raw_value.strip()
    val_upper = val.upper()

    # Rule 1: If contains known US/state keywords → mark as 'International'
    if any(x in val_upper for x in ['USA', 'U.S', 'AMERICA', 'CALIFORNIA', 'OREGON', 'HAWAII', 'NEW YORK',
                                    'TEXAS', 'ARIZONA', 'GUAM', 'WASHINGTON', 'ALASKA', 'PHILIPPINES', 'NORTH CAROLINA']):
        first_draft_mapping[raw_value] = 'International'
        continue

    # Rule 2: If ends with 2-letter US state abbreviation → 'International'
    words = val_upper.split()
    if words and words[-1] in {
        'HI', 'WA', 'CA', 'NY', 'TX', 'AZ', 'OR', 'GU', 'AK', 'NC'
    }:
        first_draft_mapping[raw_value] = 'International'
        continue

    # Rule 3: Dot-separated acronym (e.g. A.E.S → AES → match)
    clean_acronym = val_upper.replace('.', '')
    if clean_acronym in known_acronym_mapping:
        first_draft_mapping[raw_value] = known_acronym_mapping[clean_acronym]
        continue

    # Rule 4: Comma-separated location and acronym (e.g., "Ebeye, ECES")
    if ',' in val_upper:
        parts = [p.strip() for p in val_upper.split(',')]
        if len(parts) == 2:
            _, possible_acronym = parts
            possible_acronym_clean = possible_acronym.replace('.', '')
            if possible_acronym_clean in known_acronym_mapping:
                first_draft_mapping[raw_value] = known_acronym_mapping[possible_acronym_clean]
                continue

    # Rule 5: Exact acronym match (e.g., LHS)
    if val_upper in known_acronym_mapping:
        first_draft_mapping[raw_value] = known_acronym_mapping[val_upper]
        continue

    # Rule 6: Fuzzy match against official school names
    best_match = difflib.get_close_matches(val, official_school_names, n=1, cutoff=0.75)
    if best_match:
        first_draft_mapping[raw_value] = best_match[0]
    else:
        first_draft_mapping[raw_value] = None  # unresolved

# Step 4: Print for review
print("🧠 First Draft Mapping for 'Transferred FROM':")
for k in sorted(first_draft_mapping):
    print(f"'{k}' → '{first_draft_mapping[k]}'")


# %%
# Cleanup Ethnicities

# Step 1: Define initial mapping (you can expand this later)
ethnicity_mapping = {
    '-': 'Marshallese',
    'AFRIKAANS': 'Afrikaners',
    'CHINESE': 'Chinese',
    'Chinese': 'Chinese',
    'ENGLISH': 'Marshallese',
    'ENGLISH,ARSHALLESE': 'Marshallese',
    'ENGLISH,PALAUAN,MARSHALLESE': 'Marshallese',
    'ENGLISH/CHICHWA': 'Marshallese',
    'ENGLISH/CHINESE': 'Chinese',
    'ENGLISH/FIJIAN': 'Marshallese',
    'ENGLISH/FILIPINO': 'Filipino',
    'ENGLISH/KIRIBATI': 'Kiribatese',
    'ENGLISH/MARSHALLESE': 'Marshallese',
    'ENGLISH/MARSHALLESE/GILBERTESE': 'Marshallese',
    'ENGLISH/MARSHALLESE/KOREAN': 'Marshallese',
    'ENGLISH/MARSHALLESE/TUVALUAN': 'Marshallese',
    'ENGLISH/PIJIN': 'Solomon Islander',
    'ENGLISH/TAGALOG': 'Filipino',
    'ENGLISH/TAGALOG/ILOMGGO': 'Filipino',
    'ENGLISH/TUVALUAN': 'Tuvaluan',
    'ENGLISH/URDU': 'Pakistani',
    'FIJIAN': 'Fijian',
    'FIJIAN/TONGAN': 'Fijian',
    'FILIPINO': 'Filipino',
    'GERMAN/SPANISH': 'German',
    'I-KIRIBATI': 'Kiribatese',
    'JAPANASE': 'Japanese', 
    'JAPANESE': 'Japanese',
    'KIRIBATI': 'Kiribatese',
    'KIRIBATI/ENGLISH': 'Kiribatese',
    'KOREAN': 'Korean',
    'LATIN': 'Other',
    'MALAWIAN': 'Malawian',
    'MARSHALLESE': 'Marshallese',
    'Marshallese': 'Marshallese',
    'MARSHALLESE/AMERICAN': 'Marshallese',
    'MARSHALLESE/CHINESE': 'Marshallese',
    'MARSHALLESE/ENGLISH': 'Marshallese',
    'MARSHALLESE/ENGLISH/FIJIAN': 'Marshallese',
    'MARSHALLESE/FIJIAN': 'Marshallese',
    'MARSHALLESE/FILIPINO': 'Marshallese',
    'MARSHALLESE/KIRIBATI': 'Marshallese',
    'MARSHALLESE/KIWI': 'Marshallese',
    'MARSHALLESE/KOSRAEN': 'Marshallese',
    'MARSHALLESE/NGLISH': 'Marshallese',
    'MARSHALLESE/POHNPEAIN/ENGLISH': 'Marshallese',
    'MARSHALLESE/POHNPEIAN/AMERICAN-ITALIAN': 'Marshallese',
    'MARSHALLESE/POHNPEIN/JAPANESE': 'Marshallese',
    'MARSHALLESE/YAPESE': 'Marshallese',
    'NEPALI': 'Nepalese',
    'NIGERIAN (HAVSA)': 'Nigerien',
    'POHNPEI/MARSHALLESE/US': 'Marshallese',
    'ROTUMAN/FIJIAN': 'Fijian',
    'Solomon Islander': 'Solomon Islander',
    'SOLOMON ISLANDER': 'Solomon Islander',
    'SOUTH AFRICAN': 'South African',
    'TAIWANESE': 'Taiwanese',
    'TAIWANESE/MARSHALLESE': 'Marshallese',
    'TONGAN': 'Tongan',
    'TUVALU/ENGLISH/MARSHALLESE': 'Marshallese',
    'TUVALUAN': 'Tuvaluan',
    '': 'Marshallese',
    'nan': 'Marshallese',
    'NaN': 'Marshallese',
    'None': 'Marshallese',
    # Add more mappings as needed
}

# Step 2: Apply the mapping
combined_df_cleaned['CleanedEthnicity'] = combined_df_cleaned['Ethnicity'].astype(str).map(ethnicity_mapping)

# Step 3: Show unique raw values for inspection
unique_ethnicities = combined_df_cleaned['Ethnicity'].dropna().astype(str).unique()

print("📋 Unique values in 'Ethnicity':")
for val in sorted(unique_ethnicities, key=lambda x: x.lower()):
    print(f" - '{val}'")

# Step 4: Get valid ethnicity values from core_lookups
valid_ethnicities = {entry['N'] for entry in core_lookups['ethnicities']}

# Step 5: Count how many cleaned values match a valid one
ethnicity_counts = combined_df_cleaned['CleanedEthnicity'].value_counts(dropna=False)
valid_ethnicity_rows = combined_df_cleaned['CleanedEthnicity'].isin(valid_ethnicities).sum()
total_ethnicity_rows = combined_df_cleaned.shape[0]
invalid_ethnicity_rows = total_ethnicity_rows - valid_ethnicity_rows

# Step 6: Summary
print(f"\n📊 Summary of 'Ethnicity' column cleanup:")
print(f"✅ Valid 'CleanedEthnicity' values (from core_lookups['ethnicities']): {valid_ethnicity_rows}")
print(f"❌ Invalid or unmapped 'CleanedEthnicity' values: {invalid_ethnicity_rows}")
print(f"📋 Frequency breakdown:")
print(ethnicity_counts)


# %%
# Clean citizenship

# Step 1: Define initial mapping for common cases and typos
citizenship_mapping = {
    '-': 'Marshall Islands',
    'AMERICAN': 'USA',
    'AUSTRALIAN/KIWI': 'Australia',
    'BLACK AFRICANS': np.nan,
    'CHINA': 'China',
    'CHINESE': 'China',
    'EUROPEAN': np.nan,
    'FIJI': 'Fiji',
    'FIJIAN': 'Fiji',
    'FIJIAN/TONGAN': 'Fiji',
    'FILILPINO': 'Philippines',
    'FILIPINO': 'Philippines',
    'I-KIRIBATI': 'Kiribati',
    'ITALIAN': 'Italy',
    'JAPAN': 'Japan',
    'JAPANESE': 'Japan',
    'KIRIBATI': 'Kiribati',
    'KIRIBATI/SOLOMON ISLANDS': 'Kiribati',
    'KIRIBATI/TUVALU': 'Kiribati',
    'KOREA': 'South Korea',
    'MALAWI': 'Malawi',
    'MALAWIAN': 'Malawi',
    'MARSHALESE': 'Marshall Islands',
    'MARSHALESE/FIJIAN': 'Marshall Islands',
    'MARSHALLESE': 'Marshall Islands',
    'MARSHALLESE/AMERICAN': 'Marshall Islands',
    'MARSHALLESE/AMERICAN/CHUUKESE/JAPANESE/HAWAIIN': 'Marshall Islands',
    'MARSHALLESE/AMERICAN/JAPANESE': 'Marshall Islands',
    'MARSHALLESE/CHUUKESE,FILIPINO,YAPESE': 'Marshall Islands',
    'MARSHALLESE/FIJIAN': 'Marshall Islands',
    'MARSHALLESE/FILIPINA': 'Marshall Islands',
    'MARSHALLESE/FILIPINO': 'Marshall Islands',
    'MARSHALLESE/HAWAIIAN': 'Marshall Islands',
    'MARSHALLESE/JAPANESE/ITALIAN': 'Marshall Islands',
    'MARSHALLESE/KIRIBATI': 'Marshall Islands',
    'MARSHALLESE/KIRIBATI/TUVALU': 'Marshall Islands',
    'MARSHALLESE/KIRIBATI/TUVALUAN': 'Marshall Islands',
    'MARSHALLESE/KIWI': 'Marshall Islands',
    'MARSHALLESE/KOREAN': 'Marshall Islands',
    'MARSHALLESE/LATINA': 'Marshall Islands',
    'MARSHALLESE/NZ': 'Marshall Islands',
    'MARSHALLESE/PALAUAN': 'Marshall Islands',
    'MARSHALLESE/PALAUAN,GERMAN,CHAMORRO,KOREAN,FILIPINO,PALAU': 'Marshall Islands',
    'MARSHALLESE/POHNPEAN': 'Marshall Islands',
    'MARSHALLESE/POHNPEI': 'Marshall Islands',
    'MARSHALLESE/POHNPEIAN/AMERICAN/ITALIAN': 'Marshall Islands',
    'MARSHALLESE/SAMOAN': 'Marshall Islands',
    'MARSHALLESE/TAIWANESE': 'Marshall Islands',
    'MARSHALLESE/YAPESE': 'Marshall Islands',
    'MMARSHALLESE/KIWI': 'Marshall Islands',
    'NEPALI': 'Nepal',
    'NIGERIA': 'Niger',
    'PACIFIC ISLANDER': 'Other Pacific Island',
    'PAKISTANI': 'Pakistan',
    'PAPUA NEW GUINEA': 'Papua NEw Guinea',
    'PHP': 'Philippines',
    'RMI': 'Marshall Islands',
    'RMI/NZ': 'Marshall Islands',
    'RMI/USA': 'Marshall Islands',
    'SOLOMON ISLANDER': 'Solomon Islands',
    'SOLOMON ISLANDER/I-KIRIBATI': 'Solomon Islands',
    'SOLOMON ISLANDS': 'Solomon Islands',
    'SOUTH AFRICAN': np.nan,
    'TONGAN': 'Tonga',
    'TUVALU': 'Tuvalu',
    'TUVALUAN': 'Tuvalu',
    'US': 'USA',
    'USA': 'USA',
    'USA/RMI': 'Marshall Islands',
    '': 'Marshall Islands',
    'None': 'Marshall Islands',
    'NaN': 'Marshall Islands',
    'nan': 'Marshall Islands',
    # Add more as needed
}

# Step 2: Apply mapping to a new standardized column
combined_df_cleaned['CleanedCitizenship'] = combined_df_cleaned['Citizenship'].astype(str).map(citizenship_mapping)

# Step 3: Show unique raw values in original column
unique_citizenship_values = combined_df_cleaned['Citizenship'].dropna().astype(str).unique()

print("📋 Unique values in 'Citizenship':")
for val in sorted(unique_citizenship_values, key=lambda x: x.lower()):
    print(f" - '{val}'")

# Step 4: Get valid official values from core_lookups['citizenships']
valid_citizenships = {entry['N'] for entry in core_lookups['nationalities']}

# Step 5: Count how many rows match valid citizenships
citizenship_counts = combined_df_cleaned['CleanedCitizenship'].value_counts(dropna=False)
valid_citizenship_rows = combined_df_cleaned['CleanedCitizenship'].isin(valid_citizenships).sum()
total_citizenship_rows = combined_df_cleaned.shape[0]
invalid_citizenship_rows = total_citizenship_rows - valid_citizenship_rows

# Step 6: Summary
print(f"\n📊 Summary of 'Citizenship' column cleanup:")
print(f"✅ Valid 'CleanedCitizenship' values (from core_lookups['citizenships']): {valid_citizenship_rows}")
print(f"❌ Invalid or unmapped 'CleanedCitizenship' values: {invalid_citizenship_rows}")
print(f"📋 Frequency breakdown:")
print(citizenship_counts)

# Step 7: Show a sample of invalid rows
invalid_citizenship_df = combined_df_cleaned[
    ~combined_df_cleaned['CleanedCitizenship'].isin(valid_citizenships)
]

print("\n🚫 Sample rows with invalid or unmapped 'Citizenship' values:")
display(invalid_citizenship_df[['SchoolName', 'FirstName', 'LastName', 'Citizenship', 'CleanedCitizenship']].head(10))


# %%
# Clean 'Special Education Student'

# Step 1: Define mapping
sped_mapping = {
    'SPED': 'Yes',
    'YES': 'Yes',
    'Yes': 'Yes',
    '': np.nan,
    'None': np.nan,
    'nan': np.nan,
    'NaN': np.nan
    # Add more mappings if needed
}

# Step 2: Apply mapping to a new standardized column
combined_df_cleaned['CleanedSpEdStudent'] = combined_df_cleaned['Special\n Education\n Student'].astype(str).str.strip().map(sped_mapping)

# Step 3: Get valid official values (in this case only 'Yes' is valid)
valid_sped_values = {'Yes'}

# Step 4: Count how many rows match valid values
sped_counts = combined_df_cleaned['CleanedSpEdStudent'].value_counts(dropna=False)
valid_sped_rows = combined_df_cleaned['CleanedSpEdStudent'].isin(valid_sped_values).sum()
total_sped_rows = combined_df_cleaned.shape[0]
invalid_sped_rows = total_sped_rows - valid_sped_rows

# Step 5: Summary
print(f"\n📊 Summary of 'Special Education Student' column cleanup:")
print(f"✅ Valid 'CleanedSpEdStudent' values (only 'Yes'): {valid_sped_rows}")
print(f"❌ Invalid or unmapped 'CleanedSpEdStudent' values: {invalid_sped_rows}")
print(f"📋 Frequency breakdown:")
print(sped_counts)

# Step 6: Show a sample of invalid rows
invalid_sped_df = combined_df_cleaned[
    ~combined_df_cleaned['CleanedSpEdStudent'].isin(valid_sped_values)
]

print("\n🚫 Sample rows with invalid or unmapped 'Special Education Student' values:")
display(invalid_sped_df[['SchoolName', 'FirstName', 'LastName', 'Special\n Education\n Student', 'CleanedSpEdStudent']].head(10))


# %%
# # Backport data from MIEMIS cleaned up records

# # Step 1: Compute the match_key in both dataframes (if not already done)
# combined_df_cleaned['match_key'] = (
#     combined_df_cleaned['FirstName'].str.strip().str.upper() + '|' +
#     combined_df_cleaned['LastName'].str.strip().str.upper() + '|' +
#     combined_df_cleaned['FixedParsedBirthDate'].astype(str)
# )

# df_enrolments['match_key'] = (
#     df_enrolments['stuGiven'].str.strip().str.upper() + '|' +
#     df_enrolments['stuFamilyName'].str.strip().str.upper() + '|' +
#     df_enrolments['stuDoB'].astype(str)
# )

# # Step 2: Build lookup
# studentid_lookup = df_enrolments.set_index('match_key')['stuCardID'].to_dict()

# # Step 3: Backport with tracking
# def backport_student_id(row):
#     original = row.get('StudentID')
#     if pd.isna(original) or str(original).strip() == '':
#         return studentid_lookup.get(row['match_key'], original)
#     return original

# # Store old values to compare
# before = combined_df_cleaned['StudentID'].copy()

# # Apply update
# combined_df_cleaned['StudentID'] = combined_df_cleaned.apply(backport_student_id, axis=1)

# # Step 4: Compare before and after
# had_id_before = before.notna() & (before.astype(str).str.strip() != '')
# has_id_after = combined_df_cleaned['StudentID'].notna() & (combined_df_cleaned['StudentID'].astype(str).str.strip() != '')
# updated_count = has_id_after & ~had_id_before

# print(f"🆔 StudentID backport complete.")
# print(f"🔹 Had StudentID before: {had_id_before.sum()}")
# print(f"🔹 Has StudentID now:    {has_id_after.sum()}")
# print(f"🔄 Rows updated from df_enrolments: {updated_count.sum()} / {len(combined_df_cleaned)}")



# %%
# Backport a whole bunch of other data from MIEMIS

# Step 1: Build match key in both DataFrames
combined_df_cleaned['match_key'] = (
    combined_df_cleaned['FirstName'].str.strip().str.upper() + '|' +
    combined_df_cleaned['LastName'].str.strip().str.upper() + '|' +
    combined_df_cleaned['FixedParsedBirthDate'].astype(str)
)

df_enrolments['match_key'] = (
    df_enrolments['stuGiven'].str.strip().str.upper() + '|' +
    df_enrolments['stuFamilyName'].str.strip().str.upper() + '|' +
    df_enrolments['stuDoB'].astype(str)
)

# Step 2: Drop duplicates in df_enrolments by match_key (keep last)
df_enrolments_deduped = df_enrolments.drop_duplicates(subset='match_key', keep='last')

# Step 3: Convert enrolments to dictionary keyed by match_key
enrolments_dict = df_enrolments_deduped.set_index('match_key').to_dict(orient='index')

# Step 4: Define mappings from df_enrolments → combined_df_cleaned
field_map = {
    'stuCardID': 'StudentID',
    'stuEthnicity': 'CleanedEthnicity',
    'stueSpEdStr': 'CleanedSpEdStudent',
    'SpEdEnv': 'IDEA School Age',
    'SpEdDis': 'Disability',
    'SpEdEng': 'English Learner',
    'stueSpEdHasAccomodationStr': 'Has SBA Accommodation',
    'SpEdAcc': 'Type of Accommodation',
    'SpEdAss': 'Assessment Type',
}

# Step 5: Ensure all target columns exist in combined_df_cleaned
for tgt_field in field_map.values():
    if tgt_field not in combined_df_cleaned.columns:
        combined_df_cleaned[tgt_field] = pd.NA

# Step 6: Initialize stats and capture pre-update counts
backport_stats = {v: 0 for v in field_map.values()}
pre_backport_counts = {
    v: combined_df_cleaned[v].notna().sum() for v in field_map.values()
}

# Step 7: Apply backport logic row-by-row
def apply_backport(row):
    record = enrolments_dict.get(row['match_key'])
    if not record:
        return row  # No match

    for src_field, tgt_field in field_map.items():        
        if pd.isna(row[tgt_field]) or str(row[tgt_field]).strip() == '':
            if src_field == 'SpEdEnv' and record.get('stueClass', '').upper() in {'GK', 'GPREK'}:
                continue
            value = record.get(src_field)
            if pd.notna(value) and str(value).strip() != '':
                row[tgt_field] = value
                backport_stats[tgt_field] += 1
    return row

combined_df_cleaned = combined_df_cleaned.apply(apply_backport, axis=1)

# Step 8: Report
print("📋 Backport Summary:")
for field in field_map.values():
    before = pre_backport_counts.get(field, 0)
    after = combined_df_cleaned[field].notna().sum()
    updated = backport_stats[field]
    print(f"🔹 {field}: {before} → {after} (newly updated: {updated})")


# %%
# Fill in missing StudentID using predictable format: hash of FIRSTNAME|LASTNAME|DOB
import hashlib

def generate_student_id(row):
    first = str(row['FirstName']).strip().upper()
    last = str(row['LastName']).strip().upper()
    dob = str(row['FixedParsedBirthDate'])
    base_str = f"{first}|{last}|{dob}"
    hash_obj = hashlib.md5(base_str.encode('utf-8'))
    short_hash = hash_obj.hexdigest()[:12]
    return f"SID{short_hash}"

# Identify missing StudentIDs
missing_mask = combined_df_cleaned['StudentID'].isna() | (combined_df_cleaned['StudentID'].astype(str).str.strip() == '')

# Fill in only the missing values
combined_df_cleaned.loc[missing_mask, 'StudentID'] = combined_df_cleaned[missing_mask].apply(generate_student_id, axis=1)

print(f"✅ Filled {missing_mask.sum()} missing StudentID values using predictable hash-based format.")


# %%
# Prepare a draft column mapping

# Load the Excel workbook and the Students sheet
new_workbook_path = os.path.join(output_directory, empty_census_workbook_filename)
wb = load_workbook(filename=new_workbook_path, data_only=True)
ws = wb['Students']

# Extract column headers from row 3
excel_headers = [cell.value for cell in next(ws.iter_rows(min_row=3, max_row=3)) if cell.value]

# Lowercase and clean versions of Excel headers for fuzzy matching
excel_headers_clean = [str(h).strip().lower().replace(' ', '').replace('_', '') for h in excel_headers]

# Clean and prepare DataFrame column headers
df_columns = list(combined_df_cleaned.columns)
df_columns_clean = [str(c).strip().lower().replace(' ', '').replace('_', '') for c in df_columns]

# Attempt to match by index and generate draft mapping
mapping = {}
for df_col, df_col_clean in zip(df_columns, df_columns_clean):
    best_match = None
    for excel_col, excel_col_clean in zip(excel_headers, excel_headers_clean):
        if df_col_clean == excel_col_clean:
            best_match = excel_col
            break
        if df_col_clean in excel_col_clean or excel_col_clean in df_col_clean:
            best_match = excel_col
    if best_match:
        mapping[df_col] = best_match

# Print the mapping string
print("# Draft column mapping (cleaned_df → Excel Students sheet)\ncolumn_mapping = {")
for k, v in mapping.items():
    print(f"    '{k}': '{v}',")
print("}")


# %%
# Populate 'First\nLanguage' based on 'CleanedEthnicity' containing 'Marshallese'

mask = combined_df_cleaned['CleanedEthnicity'].astype(str).str.contains('Marshallese', case=False, na=False)
combined_df_cleaned.loc[mask, 'First\n Language'] = 'Marshallese'

print(f"✅ Set 'First\\nLanguage' to 'Marshallese' for {mask.sum()} students based on ethnicity.")


# %%
# %%time
import xlwings as xw
import pandas as pd

# 🔁 Always start with a clean copy of the workbook
new_workbook_path = os.path.join(output_directory, empty_census_workbook_filename)
clean_workbook_path = os.path.join(output_directory, clean_census_workbook_filename)

# Delete the previous version if it exists
if os.path.exists(new_workbook_path):
    try:
        os.remove(new_workbook_path)
        print(f"🗑️ Removed previous workbook: {new_workbook_path}")
    except Exception as e:
        raise RuntimeError(f"Failed to delete existing workbook: {new_workbook_path}\n{e}")

# Copy the clean template
try:
    shutil.copyfile(clean_workbook_path, new_workbook_path)
    print(f"📄 Copied clean template to: {new_workbook_path}")
except Exception as e:
    raise RuntimeError(f"Failed to copy clean workbook template.\n{e}")

# Mapping from cleaned dataframe to census workbook dataframe columns (edit as needed)
column_mapping = {
    'CleanedSchoolName': 'School Name',
    'SchoolYear': 'SchoolYear',
    'StudentID': 'National Student ID',
    'FirstName': 'First Name',
    'LastName': 'Last Name',
    'Gender': 'Gender',
    'Assessment Type': 'Assessment Type',
    'Exiting': 'Exiting',
    'First\n Language': 'Language',
    'FixedParsedBirthDate': 'Date of Birth',
    'DoB Estimate': 'DoB Estimate',
    'InferredAttendedECE': 'Attended ECE',
    'Grade': 'Grade Level',
    'CleanedFrom': 'From',
    'InferredTransferredFromWhichSchool': 'Transferred From which school',
    #'CleanedTransferredInDate': 'Transfer In Date',
    'CleanedEthnicity': 'Ethnicity',
    'CleanedCitizenship': 'Citizenship',
    'CleanedSpEdStudent': 'SpEd Student',
    # 'IDEA ECE': 'IDEA ECE',
    'IDEA School Age': 'IDEA School Age',
    'Disability': 'Disability',
    'English Learner': 'English Learner',
    'Has IEP': 'Has IEP',
    'Has SBA Accommodation': 'Has SBA Accommodation',
    'Type of Accommodation': 'Type of Accommodation',
    # 'Assessment Type'
    # 'Exiting'
    # 'Exiting Date'
    # 'Days Absent'
    # 'Completed?'
    # 'Outcome'
    # 'Dropout Reason'
    # 'Expulsion Reason'
    # 'Transferred To Which School'
    # 'Post-secondary study'
    # 'Bullied'
}
    
# Open the workbook with xlwings (preserves formatting, macros, formulas)
try:
    app = xw.App(visible=False)
    wb = app.books.open(new_workbook_path)
    
    # Get the correct worksheet
    sheet_name = 'Students'
    ws = wb.sheets[sheet_name]
    
    # ✅ Unprotect the correct sheet
    ws.api.Unprotect()  # Add password if needed
    
    # Prepare DataFrame
    df_to_insert = combined_df_cleaned[list(column_mapping.keys())].copy()
    df_to_insert.rename(columns=column_mapping, inplace=True)
    df_to_insert = df_to_insert.astype(object).where(pd.notna(df_to_insert), None)
    
    # Read Excel headers from row 3
    header_row = 3
    excel_headers = ws.range((header_row, 1)).expand('right').value
    header_indices = {
        header: idx + 1 for idx, header in enumerate(excel_headers) if header in df_to_insert.columns
    }
    
    # Test on small subset
    #df_to_insert = df_to_insert[:1000].copy()
    
    # # Write row by row approach
    # # Start inserting data from row 4
    # start_row = header_row + 1
    # for i, (_, row) in enumerate(df_to_insert.iterrows(), start=start_row):
    #     #print(f"Writing row {i}")
    #     #if (i - start_row + 1) % 50 == 0:
    #     #    print(f"Writing row {i}")
    #     for col_name, value in row.items():
    #         col_idx = header_indices.get(col_name)
    #         if col_idx:
    #             ws.cells(i, col_idx).value = value
    
    # Write using diagnostic vectorized approach
    start_row = header_row + 1
    num_rows = len(df_to_insert)
    
    invalid_columns = []
    
    for col_name, col_idx in header_indices.items():
        col_values = df_to_insert[col_name].tolist()
        try:
            # Try writing one column vector at a time
            ws.range((start_row, col_idx), (start_row + num_rows - 1, col_idx)).value = [[v] for v in col_values]
        except Exception as e:
            print(f"❌ Error in column: {col_name} (Excel column {col_idx})")
            print(e)
            invalid_columns.append(col_name)

    if len(invalid_columns) > 0:
        print(f"\n🚨 Columns that failed to write: {invalid_columns}")

    
    # Optionally re-protect the sheet
    # ws.api.Protect()

    # Save and close
    wb.save()
finally:
    wb.close()
    app.quit()

print("✅ Data successfully injected into Excel workbook without touching formulas or formatting.")


# %%
df_to_insert[:3]

# %%
df_to_insert.columns

# %%
# Look at duplicates

# Normalize the DataFrame
df = df_to_insert.copy()

# Clean up key fields
df['National Student ID'] = df['National Student ID'].astype(str).str.strip().str.upper()
df['First Name'] = df['First Name'].astype(str).str.strip().str.upper()
df['Last Name'] = df['Last Name'].astype(str).str.strip().str.upper()
df['Date of Birth'] = pd.to_datetime(df['Date of Birth'], errors='coerce')

# ✅ 1. Duplicates by National Student ID + First + Last + DOB
dupe_keys_1 = ['National Student ID', 'First Name', 'Last Name', 'Date of Birth']
df_to_insert_dupes1 = df[df.duplicated(dupe_keys_1, keep=False)].sort_values(by=dupe_keys_1)

# ✅ 2. Duplicates by First + Last + DOB only (ignoring NSID)
dupe_keys_2 = ['First Name', 'Last Name', 'Date of Birth']
df_to_insert_dupes2 = df[df.duplicated(dupe_keys_2, keep=False)].sort_values(by=dupe_keys_2)

# ✅ Summary
print(f"🧾 Duplicates by National Student ID + Name + DOB: {df_to_insert_dupes1.shape[0]} rows in {df_to_insert_dupes1[dupe_keys_1].drop_duplicates().shape[0]} groups.")
print(f"🧾 Duplicates by Name + DOB only: {df_to_insert_dupes2.shape[0]} rows in {df_to_insert_dupes2[dupe_keys_2].drop_duplicates().shape[0]} groups.")


# %%
