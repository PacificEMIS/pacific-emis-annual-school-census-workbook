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
    return config['output_directory'], config['source_workbook_filename'], config['empty_census_workbook_filename'], config['clean_census_workbook_filename'], config['delete_census_workbook_filename']
    

# Test loading configuration
output_directory, source_workbook_filename, empty_census_workbook_filename, clean_census_workbook_filename, delete_census_workbook_filename = load_config()
print("Configuration loaded successfully.")

# Import some lookups
# %store -r core_lookups student_lookups censusworkbook_lookups
# %store -r all_teachers all_schools 
# %store -r df_teacher_recent_survey_data

# %%
# Load workbook

# Combine the directory and filename
workbook_path = os.path.join(output_directory, source_workbook_filename)

# Load the workbook (but not yet any sheet into memory)
xls = pd.ExcelFile(workbook_path)

# See the available sheet names
print("Sheets found:", xls.sheet_names)


# %%
# ✅ Get Staff data from "Staff Roster" sheet

# Build the full path
workbook_path = os.path.join(output_directory, source_workbook_filename)

# Load the Excel file
xls = pd.ExcelFile(workbook_path)

# Specify the staff sheet
staff_sheet = "Staff Roster"

# Load the sheet
df_staff_raw = xls.parse(staff_sheet)

# Display basic preview
display(df_staff_raw.head(3))

# Summary info
total_rows = df_staff_raw.shape[0]
column_list = list(df_staff_raw.columns)

print(f"✅ Successfully loaded '{staff_sheet}' with {total_rows} rows.")
print(f"🧾 Columns in df_staff_raw ({len(column_list)}):")
print(column_list)

# Summary per School (if present)
if 'School' in df_staff_raw.columns:
    print("\n📊 Row counts per School:")
    print(df_staff_raw['School'].value_counts(dropna=False).sort_index())
else:
    print("ℹ️ 'School' column not found in the staff roster.")


# %%
# ✅ Filter staff records for the current school year 'SY24-25'

# Define the target year code
target_staff_year = "SY24-25"

# Filter records
df_staff_filtered = df_staff_raw[df_staff_raw['Year'] == target_staff_year].copy()

# Display a preview
display(df_staff_filtered.head(3))

# Summary info
filtered_count = df_staff_filtered.shape[0]
print(f"✅ Filtered staff records to {filtered_count} row(s) for Year == '{target_staff_year}'.")


# %%
# ✅ Remove junk rows from staff data (missing key info)

# Define mask: True for rows where all key fields are either NaN or blank
junk_mask = (
    df_staff_filtered['First Name'].isna() | df_staff_filtered['First Name'].astype(str).str.strip().eq('')
) & (
    df_staff_filtered['Last Name'].isna() | df_staff_filtered['Last Name'].astype(str).str.strip().eq('')
) & (
    df_staff_filtered['School'].isna() | df_staff_filtered['School'].astype(str).str.strip().eq('')
)

# Count and drop junk rows
junk_count = junk_mask.sum()
df_staff_filtered = df_staff_filtered[~junk_mask].copy()

# Display preview
display(df_staff_filtered.head(3))

# Summary
print(f"🧹 Removed {junk_count} junk row(s) with missing First Name, Last Name, and School.")
print(f"✅ Remaining valid staff records: {df_staff_filtered.shape[0]}")


# %%
# ✅ Clean 'Year' column into 'CleanedYear' with full 4-digit format

import re

def standardize_year(year_value):
    if isinstance(year_value, str):
        match = re.match(r'^SY?(\d{2})[-–](\d{2})$', year_value.strip())
        if match:
            start, end = match.groups()
            start_full = int(start)
            end_full = int(end)
            # Assume school years are within the same century (20xx)
            if start_full < 50:  # e.g., 24 -> 2024
                start_full += 2000
            else:  # unlikely but handles 90s as 199x
                start_full += 1900
            if end_full < 50:
                end_full += 2000
            else:
                end_full += 1900
            return f"SY{start_full}-{end_full}"
    return year_value  # fallback to original if no match

# Apply the cleaning
df_staff_filtered['CleanedYear'] = df_staff_filtered['Year'].apply(standardize_year)

# Show unique cleaned values
unique_cleaned_years = df_staff_filtered['CleanedYear'].dropna().unique()
print("📅 Unique cleaned 'CleanedYear' values:")
for val in sorted(unique_cleaned_years):
    print(f" - {val}")


# %%
# ✅ Clean 'School' into 'CleanedSchool' using lookup with manual mapping support

# Step 1: Extract valid school names from lookup
valid_school_names = {entry['N'] for entry in censusworkbook_lookups['schoolCodes']}

# Step 2: Create initial CleanedSchool where value is valid
df_staff_filtered['CleanedSchool'] = df_staff_filtered['School'].where(
    df_staff_filtered['School'].isin(valid_school_names), pd.NA
)

# Step 3: Identify unmatched values to build mapping
unmatched_schools = df_staff_filtered.loc[df_staff_filtered['CleanedSchool'].isna(), 'School'].dropna().unique()

print("📋 Unmatched 'School' values (please review and build mapping):")
for val in sorted(unmatched_schools):
    print(f" - '{val}'")

# Step 4: Create manual mapping dictionary (🔁 you update this)
manual_school_mapping = {
    'Jabenoden ES-Jaluit': 'Jabnodren Elementary School',
    # Example:
    # 'Ajeltake Elem School': 'Ajeltake Elementary School',
}

# Step 5: Apply manual mapping to fill in CleanedSchool
df_staff_filtered.loc[
    df_staff_filtered['School'].isin(manual_school_mapping.keys()),
    'CleanedSchool'
] = df_staff_filtered['School'].map(manual_school_mapping)

# Step 6: Final validation and summary
final_valid_count = df_staff_filtered['CleanedSchool'].notna().sum()
final_invalid_count = df_staff_filtered.shape[0] - final_valid_count
still_unmatched = df_staff_filtered.loc[df_staff_filtered['CleanedSchool'].isna(), 'School'].dropna().unique()

print(f"\n🏁 Final CleanedSchool summary:")
print(f"✅ Matched or corrected: {final_valid_count}")
print(f"❌ Still unmatched: {final_invalid_count}")
if still_unmatched.size > 0:
    print("⚠️ Still unmatched values after manual mapping:")
    for val in sorted(still_unmatched):
        print(f" - '{val}'")


# %%
# ✅ Build CleanedSchNo based on CleanedSchool using lookup

# Step 1: Build dictionary from CleanedSchool to School Code
school_name_to_code = {entry['N']: entry['C'] for entry in censusworkbook_lookups['schoolCodes']}

# Step 2: Map CleanedSchool to CleanedSchNo
df_staff_filtered['CleanedSchNo'] = df_staff_filtered['CleanedSchool'].map(school_name_to_code)

# Step 3: Validation summary
valid_schno_count = df_staff_filtered['CleanedSchNo'].notna().sum()
invalid_schno_rows = df_staff_filtered['CleanedSchool'].notna() & df_staff_filtered['CleanedSchNo'].isna()
unmatched_cleaned_schools = df_staff_filtered.loc[invalid_schno_rows, 'CleanedSchool'].unique()

print(f"✅ Successfully populated CleanedSchNo for {valid_schno_count} rows.")
if unmatched_cleaned_schools.size > 0:
    print(f"⚠️ Could not find CleanedSchNo for the following CleanedSchool values:")
    for name in sorted(unmatched_cleaned_schools):
        print(f" - '{name}'")


# %%
# ✅ Clean and Validate Gender in Staff Data

# Step 1: Define valid values
valid_genders = {'Male', 'Female'}

# Step 2: Standardize raw Gender values
df_staff_filtered['Gender'] = df_staff_filtered['Gender'].astype(str).str.strip()

# Step 3: Flag initial invalids
df_staff_filtered['Invalid_Gender'] = ~df_staff_filtered['Gender'].isin(valid_genders)
initial_invalids = df_staff_filtered[df_staff_filtered['Invalid_Gender']]
print(f"⚠️ Found {initial_invalids.shape[0]} row(s) with invalid Gender before correction.")

# Step 4: Apply corrections
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
df_staff_filtered['CleanedGender'] = df_staff_filtered['Gender'].replace(gender_corrections)

# Step 5: Re-flag invalids after correction
df_staff_filtered['Invalid_Gender'] = ~df_staff_filtered['CleanedGender'].isin(valid_genders)

# Step 6: Summary
remaining_invalids = df_staff_filtered['Invalid_Gender'].sum()
print(f"✅ Gender correction complete. Remaining invalid rows: {remaining_invalids}")

# Optional: Preview a few remaining invalids
if remaining_invalids:
    display(df_staff_filtered[df_staff_filtered['Invalid_Gender']][['School', 'CleanedGender', 'SourceSheet']].head(10))


# %%
# ✅ Clean and Parse 'Date of Birth' into 'CleanedDateofBirth'

from dateutil import parser
import datetime

def try_parse_dob(value):
    """Attempt to parse a date of birth value with fallback logic."""
    if pd.isna(value) or str(value).strip() == '':
        return None

    # Handle Excel serial numbers
    if isinstance(value, (int, float)) and value > 59:
        try:
            return pd.to_datetime('1899-12-30') + pd.to_timedelta(int(value), unit='D')
        except Exception:
            return None

    # Try parsing string formats
    str_val = str(value).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%B %d, %Y", "%d-%b-%y"):
        try:
            return datetime.datetime.strptime(str_val, fmt).date()
        except Exception:
            continue

    # Try auto parser
    try:
        return parser.parse(str_val, dayfirst=False).date()
    except Exception:
        return None

# Step 1: Apply DOB parser
df_staff_filtered['CleanedDateofBirth'] = df_staff_filtered['Date of Birth'].apply(try_parse_dob)

# Step 2: Flag rows that failed parsing
df_staff_filtered['Invalid_DOB'] = df_staff_filtered['CleanedDateofBirth'].isna() & df_staff_filtered['Date of Birth'].notna()

# Step 3: Summary
total_invalid = df_staff_filtered['Invalid_DOB'].sum()
print(f"✅ Completed Date of Birth cleaning.")
print(f"⚠️ Found {total_invalid} row(s) with unparseable Date of Birth.")

# Optional: Show a few problematic rows
if total_invalid:
    display(df_staff_filtered[df_staff_filtered['Invalid_DOB']][['School', 'First Name', 'Last Name', 'Date of Birth']].head(10))


# %%
from datetime import datetime

# Step 1: Build a lookup dictionary from existing teachers based on (First, Last) name
teacher_dob_lookup = {
    (t['tGiven'].strip().lower(), t['tSurname'].strip().lower()): t['tDOB']
    for t in all_teachers
    if t['tGiven'] and t['tSurname'] and t['tDOB'] not in (None, '', '1900-01-01T00:00:00')  # ignore placeholders
}

# Step 2: Initialize counters
forward_ported_missing = 0
forward_ported_unparsed = 0

# Step 3: Go through rows and attempt forward porting
for idx, row in df_staff_filtered.iterrows():
    dob = row.get('CleanedDateofBirth')
    first = str(row.get('First Name', '')).strip().lower()
    last = str(row.get('Last Name', '')).strip().lower()

    if pd.isna(dob) and (first, last) in teacher_dob_lookup:
        dob_str = teacher_dob_lookup[(first, last)]
        try:
            parsed_dob = parser.parse(dob_str).date()
            df_staff_filtered.at[idx, 'CleanedDateofBirth'] = parsed_dob

            # Classify forward-port type
            original_raw = str(row.get('Date of Birth', '')).strip()
            if original_raw == '' or pd.isna(row['Date of Birth']):
                forward_ported_missing += 1
            else:
                forward_ported_unparsed += 1

        except Exception:
            continue  # if even the original tDOB is unparseable, skip

# Step 4: Print summary
print(f"📦 Forward-port complete.")
print(f"✅ {forward_ported_missing} record(s) filled from teachers where DOB was missing.")
print(f"♻️  {forward_ported_unparsed} record(s) filled from teachers where DOB was unparsed.")



# %%
from difflib import get_close_matches

# Step 1: Load valid citizenship names from lookup
valid_citizenships = {entry['N'] for entry in censusworkbook_lookups['nationalities']}
valid_citizenships_lower = {v.lower(): v for v in valid_citizenships}

# Step 2: Extract unique raw citizenships
raw_citizenships = df_staff_filtered['Citizenship'].dropna().astype(str).str.strip().unique()

# Step 3: Build draft mapping
draft_mapping = {}

for raw_value in raw_citizenships:
    val = raw_value.strip()
    val_lower = val.lower()

    # 1. Exact match
    if val in valid_citizenships:
        draft_mapping[val] = val
        continue

    # 2. Case-insensitive match
    if val_lower in valid_citizenships_lower:
        draft_mapping[val] = valid_citizenships_lower[val_lower]
        continue

    # 3. Fuzzy match
    close = get_close_matches(val, valid_citizenships, n=1, cutoff=0.8)
    if close:
        draft_mapping[val] = close[0]
    else:
        draft_mapping[val] = None  # Needs manual review

# Step 4: Apply mapping
df_staff_filtered['CleanedCitizenship'] = (
    df_staff_filtered['Citizenship'].astype(str).str.strip().map(draft_mapping)
)

# Step 5: Summary
print("📋 Draft automatic mapping:")
for k, v in sorted(draft_mapping.items()):
    print(f"'{k}': '{v}'")

valid_count = df_staff_filtered['CleanedCitizenship'].isin(valid_citizenships).sum()
total_rows = df_staff_filtered.shape[0]
invalid_count = total_rows - valid_count

print(f"\n✅ Auto-mapped valid citizenships: {valid_count}")
print(f"❌ Remaining unmapped or invalid citizenships: {invalid_count}")


# %%
from difflib import get_close_matches

# Step 1: Load valid citizenships from lookup
valid_citizenships = {entry['N'] for entry in core_lookups['nationalities']}
valid_citizenships_lower = {v.lower(): v for v in valid_citizenships}

# Step 2: Start with manual mapping
citizenship_mapping = {
    'Brazil': 'Brazil',
    'Cameroon': 'Cameroon',
    'Canada': 'Canada',
    'Chuuk': 'FSM',
    'Chuuk Islands': 'FSM',
    'Denmark': 'Denmark',
    'FIJ': 'Fiji',
    'FIJI': 'Fiji',
    'FSM': 'FSM',
    'Fiji': 'Fiji',
    'Fijian': 'Fiji',
    'JAPAN': 'Japan',
    'KIRIBATI': 'Kiribati',
    'Kirbati': 'Kiribati',
    'Kiribati': 'Kiribati',
    'Kiribati Isalnds': 'Kiribati',
    'Kiribati Islands': 'Kiribati',
    'Marshall Isalnds': 'Marshall Islands',
    'Marshall Islands': 'Marshall Islands',
    'NIGERIA': 'Nigeria',
    'PHI': 'Philippines',
    'PI': 'Philippines',
    'PNG': 'Papua New Guinea',
    'POH': 'FSM',
    'Pakistan': 'Pakistan',
    'Papua New Guinea': 'Papua New Guinea',
    'Philippines': 'Philippines',
    'Phillipines': 'Philippines',
    'Pohnpei': 'FSM',
    'RMI': 'Marshall Islands',
    'ROP': 'Palau',
    'ROV': 'Vanuatu',
    'SOL': 'Solomon Islands',
    'SOLOMON': 'Solomon Islands',
    'Solmon Islands': 'Solomon Islands',
    'Solomon Islands': 'Solomon Islands',
    'South Africa': 'South Africa',
    'TUV': 'Tuvalu',
    'Taiwan': 'Taiwan',
    'Tuvalu': 'Tuvalu',
    'UK': 'United Kingdom',
    'US': 'USA',
    'USA': 'USA',
    'Zimbabwe': 'Zimbabwe',
}

# Step 3: Get all raw values from the data
raw_citizenships = df_staff_filtered['Citizenship'].dropna().astype(str).str.strip().unique()

# Step 4: Try to auto-map values not already manually mapped
for val in raw_citizenships:
    if val in citizenship_mapping:
        continue  # already mapped manually

    val_stripped = val.strip()
    val_lower = val_stripped.lower()

    # Case-insensitive match
    if val_lower in valid_citizenships_lower:
        citizenship_mapping[val] = valid_citizenships_lower[val_lower]
    else:
        # Fuzzy match
        close = get_close_matches(val_stripped, valid_citizenships, n=1, cutoff=0.8)
        if close:
            citizenship_mapping[val] = close[0]
        else:
            citizenship_mapping[val] = None  # needs manual review

# Step 5: Apply mapping
df_staff_filtered['CleanedCitizenship'] = (
    df_staff_filtered['Citizenship'].astype(str).str.strip().map(citizenship_mapping)
)

# Step 6: Summary
print("\n📋 Final citizenship mapping:")
for k, v in sorted(citizenship_mapping.items()):
    print(f"'{k}': '{v}'")

valid_count = df_staff_filtered['CleanedCitizenship'].isin(valid_citizenships).sum()
invalid_count = df_staff_filtered.shape[0] - valid_count

print(f"\n✅ Auto-mapped valid citizenships: {valid_count}")
print(f"❌ Remaining unmapped or invalid: {invalid_count}")

# Optional: Show sample of invalids
invalid_df = df_staff_filtered[~df_staff_filtered['CleanedCitizenship'].isin(valid_citizenships)]
display(invalid_df[['School', 'First Name', 'Last Name', 'Citizenship', 'CleanedCitizenship']].head(10))


# %%
from difflib import get_close_matches

# Step 1: Load valid ethnicity names from lookup
valid_ethnicities = {entry['N'] for entry in censusworkbook_lookups['ethnicities']}
valid_ethnicities_lower = {v.lower(): v for v in valid_ethnicities}

# Step 2: Extract unique raw ethnicities
raw_ethnicities = df_staff_filtered['Ethnicity'].dropna().astype(str).str.strip().unique()

# Step 3: Start with manual mapping
ethnicity_mapping = {
    'African': 'Other African',
    'American': np.nan,
    'Asian': 'Other Asian',
    'Brazilian': 'Brazilian',
    'British': 'Caucasian',
    'Cameroonian': 'Cameroonian',
    'Canadian': np.nan,
    'Caucasian': 'Caucasian',
    'Chuukese': 'Chuukese',
    'FIjian': 'Fijian',
    'Fijian': 'Fijian',
    'Filipino': 'Filipino',
    'Filipinos': 'Filipino',
    'Fillipino': 'Filipino',
    'I Kiribati': 'Kiribatese',
    'Japanese': 'Japanese',
    'Kiribatese': 'Kiribatese',
    'Marshallese': 'Marshallese',
    'Other': 'Other',
    'Pacific Islander': 'Other Pacific Islander',
    'Pakistanis': 'Pakistani',
    'Palauan': 'Palauan',
    'Papua New Guinea': 'Papua New Guinean',
    'Pohnpeian': 'Pohnpeian',
    'Shona': np.nan,
    'Solomon': 'Solomon Islander',
    'Taiwanese': 'Taiwanese',
    'Tuvaluan': 'Tuvaluan',
    'Vanutuan': 'Ni Vanuatu',
}

# Step 4: Add auto-mapping ONLY for values not already manually mapped
for raw_value in raw_ethnicities:
    val = raw_value.strip()
    val_lower = val.lower()

    if val in ethnicity_mapping:
        continue  # already mapped

    # 1. Exact match
    if val in valid_ethnicities:
        ethnicity_mapping[val] = val
        continue

    # 2. Case-insensitive match
    if val_lower in valid_ethnicities_lower:
        ethnicity_mapping[val] = valid_ethnicities_lower[val_lower]
        continue

    # 3. Fuzzy match
    close = get_close_matches(val, valid_ethnicities, n=1, cutoff=0.8)
    if close:
        ethnicity_mapping[val] = close[0]
    else:
        ethnicity_mapping[val] = None  # Needs manual review

# Step 5: Apply mapping
df_staff_filtered['CleanedEthnicity'] = (
    df_staff_filtered['Ethnicity'].astype(str).str.strip().map(ethnicity_mapping)
)

# Step 6: Summary
print("📋 Final ethnicity mapping:")
for k, v in sorted(ethnicity_mapping.items()):
    print(f"'{k}': '{v}'")

valid_count = df_staff_filtered['CleanedEthnicity'].isin(valid_ethnicities).sum()
total_rows = df_staff_filtered.shape[0]
invalid_count = total_rows - valid_count

print(f"\n✅ Auto-mapped valid ethnicities: {valid_count}")
print(f"❌ Remaining unmapped or invalid ethnicities: {invalid_count}")

# Step 7: Frequency breakdown
ethnicity_counts = df_staff_filtered['CleanedEthnicity'].value_counts(dropna=False)
print("\n📊 Frequency breakdown of 'CleanedEthnicity':")
print(ethnicity_counts)

# Step 8: Sample invalid rows
invalid_ethnicity_df = df_staff_filtered[
    ~df_staff_filtered['CleanedEthnicity'].isin(valid_ethnicities)
]

print("\n🚫 Sample rows with invalid or unmapped 'Ethnicity' values:")
display(invalid_ethnicity_df[['School', 'First Name', 'Last Name', 'Ethnicity', 'CleanedEthnicity']].head(100))


# %%
# Step 1: Normalize and prepare keys
df_staff_filtered['First Name'] = df_staff_filtered['First Name'].astype(str).str.strip().str.lower()
df_staff_filtered['Last Name'] = df_staff_filtered['Last Name'].astype(str).str.strip().str.lower()
df_staff_filtered['CleanedDateofBirth'] = pd.to_datetime(df_staff_filtered['CleanedDateofBirth'], errors='coerce')

# Convert to DataFrame
all_teachers_df = pd.DataFrame(all_teachers)

all_teachers_df['tGiven'] = all_teachers_df['tGiven'].astype(str).str.strip().str.lower()
all_teachers_df['tSurname'] = all_teachers_df['tSurname'].astype(str).str.strip().str.lower()
all_teachers_df['tDOB'] = pd.to_datetime(all_teachers_df['tDOB'], errors='coerce')

# Step 2: Join keys
df_staff_filtered['strict_key'] = (
    df_staff_filtered['First Name'] + '|' +
    df_staff_filtered['Last Name'] + '|' +
    df_staff_filtered['CleanedDateofBirth'].astype(str)
)

df_staff_filtered['loose_key'] = (
    df_staff_filtered['First Name'] + '|' +
    df_staff_filtered['Last Name']
)

all_teachers_df['strict_key'] = (
    all_teachers_df['tGiven'] + '|' +
    all_teachers_df['tSurname'] + '|' +
    all_teachers_df['tDOB'].astype(str)
)

all_teachers_df['loose_key'] = (
    all_teachers_df['tGiven'] + '|' +
    all_teachers_df['tSurname']
)

# Step 3: Build lookup dictionaries
strict_lookup = all_teachers_df.set_index('strict_key')['tPayroll'].to_dict()
loose_lookup = all_teachers_df.set_index('loose_key')['tPayroll'].to_dict()

# Step 4: Track initial nulls
before_filled = df_staff_filtered['RMI SS#'].isna().sum()

# Step 5: Try strict match first
df_staff_filtered['RMI SS#_filled'] = df_staff_filtered.apply(
    lambda row: strict_lookup.get(row['strict_key'], None) if pd.isna(row['RMI SS#']) else row['RMI SS#'],
    axis=1
)

# Step 6: Fallback to loose match
df_staff_filtered['RMI SS#_filled'] = df_staff_filtered.apply(
    lambda row: loose_lookup.get(row['loose_key'], row['RMI SS#_filled']) if pd.isna(row['RMI SS#']) else row['RMI SS#_filled'],
    axis=1
)

# Step 7: Finalize column
df_staff_filtered['RMI SS#'] = df_staff_filtered['RMI SS#_filled']
df_staff_filtered.drop(columns=['RMI SS#_filled', 'strict_key', 'loose_key'], inplace=True)

# Step 8: Summary
after_filled = df_staff_filtered['RMI SS#'].isna().sum()
filled_count = before_filled - after_filled

print(f"✅ RMI SS# forward-fill complete with optional DOB matching.")
print(f"📌 Initially missing: {before_filled}")
print(f"✅ Filled using strict or loose match: {filled_count}")
print(f"❌ Still missing after both passes: {after_filled}")


# %%
from difflib import get_close_matches
import numpy as np

# Step 1: Load valid qualification names and code-to-name mapping from lookup, excluding 'Academic Degree'
qual_lookup = [
    entry for entry in core_lookups['teacherQuals']
    if entry.get('G') == 'Academic Degree'
]

valid_qual_names = {entry['N'].strip() for entry in qual_lookup}
valid_qual_names_lower = {v.lower(): v for v in valid_qual_names}
qual_code_to_name = {entry['C'].strip(): entry['N'].strip() for entry in qual_lookup}

# Step 2: Extract unique raw qualifications
raw_qualifications = df_staff_filtered['Highest Qualification'].dropna().astype(str).str.strip().unique()

# Step 3: Manual mapping
qualification_mapping = {
    '2 yrs. Certificate': 'Certificate',
    'AA': 'Associate of Arts',
    'AA (Associates)': 'Associate of Arts',
    'AA In Business Management': 'Associate of Arts',
    'AA in Liberal Arts': 'Associate of Arts',
    'AS': 'Associate of Science',
    'AS (Associates)': 'Associate of Science',
    'AS CCD': 'Associate of Science',
    'AS Degree': 'Associate of Science',
    'AS Degree in Elementary Education': 'Associate of Science',
    'AS Degree in Liberal Arts': 'Associate of Science',
    'AS Education': 'Associate of Science',
    'AS In Liberal Arts': 'Associate of Science',
    'AS degree': 'Associate of Science',
    'AS degree Liber Arts': 'Associate of Science',
    'AS in Business': 'Associate of Science',
    'AS in Business/CCT': 'Associate of Science',
    'AS in Education': 'Associate of Science',
    'AS in Elem.Education': 'Associate of Science',
    'AS in Elementary Eucation': 'Associate of Science',
    'AS in Liberal Arts': 'Associate of Science',
    'AS in Liberat Arts': 'Associate of Science',
    'AS of Science': 'Associate of Science',
    'AS( Associates)': 'Associate of Science',
    'ASEE': 'Associate of Science',
    'ASEE (Associates)': 'Associate of Science',
    'ASsociate of Science': 'Associate of Science',
    'As': 'Associate of Science',
    'As Degree': 'Associate of Science',
    'Associate': 'Associate of Science',
    'Associate of Arts': 'Associate of Arts',
    'Associate of Liberal Arts': 'Associate of Arts',
    'Associate of Science': 'Associate of Science',
    'Associates': 'Associate of Science',
    'Associates/AS': 'Associate of Science',
    'B.S': 'Bachelor of Science',
    'BA': 'Bachelor of Arts',
    'BA (Bachelors)': 'Bachelor of Arts',
    'BA Elementary Education': 'Bachelor of Arts',
    'BA In Public Health': 'Bachelor of Arts',
    'BA Public Admin & Social Work': 'Bachelor of Arts',
    'BA Sociology & Social work': 'Bachelor of Arts',
    'BA degree': 'Bachelor of Arts',
    'BA in Biblical': 'Bachelor of Arts',
    'BA in Ed. Primary': 'Bachelor of Arts',
    'BA in Education': 'Bachelor of Arts',
    'BA in Theology': 'Bachelor of Arts',
    'BA of Science in Education': 'Bachelor of Arts',
    'BAEE': 'Bachelor of Arts',
    'BAEE (Bachelors)': 'Bachelor of Arts',
    'BED in Math & Physic': 'Bachelor of Education',
    'BEd': 'Bachelor of Education',
    'BS': 'Bachelor of Science',
    'BS (Bachelors)': 'Bachelor of Science',
    'BS in Edu': 'Bachelor of Science',
    'BS-Culinary': 'Bachelor of Science',
    'BSN': 'Bachelor of Science',
    'Ba  in Education': 'Bachelor of Arts',
    'Bachelor': 'Bachelor of Arts',
    'Bachelor Degree with MA unit': 'Bachelor of Arts',
    'Bachelor Science in nursing': 'Bachelor of Science',
    'Bachelor in Secondary Education (English)': 'Bachelor of Education',
    'Bachelor in Special Education': 'Bachelor of Education',
    'Bachelor of Arts': 'Bachelor of Arts',
    'Bachelor of Education': 'Bachelor of Education',
    'Bachelor of Science': 'Bachelor of Science',
    'Bachelor of Science in Business Admin': 'Bachelor of Science',
    'Bachelor of Science in Education': 'Bachelor of Science',
    'Bachelors': 'Bachelor of Arts',
    'CECE': 'Certificate ECE',
    'CECE (MOC)': 'Certificate ECE',
    'CECE Certificate': 'Certificate ECE',
    'CEDE': 'Certificate',
    'CEDE Certificate': 'Certificate',
    'CMI': 'Certificate',
    'Cert ECE': 'Certificate ECE',
    'Cert. ECE': 'Certificate ECE',
    'Certificate': 'Certificate',
    'Certificate ECE': 'Certificate ECE',
    'Certificate in Carpentry': np.nan,
    'Certificate in Organic & Agriculture': np.nan,
    'Certificate of ECE': 'Certificate ECE',
    'Certification of Completion in Teaching': 'Certification of Completion in Teaching',
    'Classroom Teacher': np.nan,
    'College': 'Certificate',
    'DECE': 'Diploma ECE',
    'DECEC': 'Diploma ECE',
    'Degree': 'Certificate',
    'Diploma': 'Diploma',
    'Diploma ECE': 'Diploma ECE',
    'Diploma ECECE': 'Diploma ECE',
    'Diploma in Education': 'Diploma in Education',
    'Diploma in Information Technology & Computer System ( Finishing up her BA in Computer System': 'Diploma',
    'Diploma in Primary Education': 'Diploma in Education',
    'Diploma of Community Social Services work APTC': 'Diploma',
    'Diploma of Education': 'Diploma in Education',
    'Diploma-IT & Networking': 'Diploma',
    'ECE Certificate': 'Certificate ECE',
    'ECEC': 'Certificate ECE',
    'ECEC Certificate': 'Certificate ECE',
    'Elem': 'Diploma in Education',
    'Elementary Diploma': 'Diploma in Education',
    'GED': 'Diploma in Education',
    'GED (Certificate)': np.nan,
    'GED (Diploma)': 'Diploma in Education',
    'GED Diploma': 'Diploma in Education',
    'H.S Diploma': 'High School',
    'HPU Cert.': np.nan,
    'HS': 'High School',
    'HS (Diploma)': 'High School',
    'HS Diploma': 'High School',
    'HS/Job Corp': 'High School',
    'High School': 'High School',
    'High School Diploma': 'High School',
    'High School diploma': 'High School',
    'High Scool Diploma': 'High School',
    'High school': 'High School',
    'Highschool': 'High School',
    'Hish School': 'High School',
    'JSMI Bible Institute': 'Certificate',
    'LA': 'Associate of Arts',
    'LA (Associates)': 'Associate of Arts',
    'Liberal Arts': 'Associate of Arts',
    'M.ED': 'Masters of Education',
    'M.ED (Masters)': 'Masters of Education',
    'M.Ed': 'Masters of Education',
    'MA': 'Masters of Arts',
    'MA (Masters)': 'Masters of Arts',
    'MA in Edu.': 'Masters of Education',
    'MBA': 'Masters of Business Administration',
    'MS': 'Masters of Science',
    'Master': 'Masters Degree',
    'Master Degree': 'Masters Degree',
    'Master Education': 'Masters of Education',
    'Master in Education': 'Masters of Education',
    'Master of Education': 'Masters of Education',
    'Master of Science': 'Masters of Science',
    'Masters': 'Masters Degree',
    'Masters In Curriculum Studies': 'Masters of Education',
    'Masters in Education': 'Masters of Education',
    'Masters of Arts': 'Masters of Arts',
    'Masters of Education': 'Masters of Education',
    'Masters of Science': 'Masters of Science',
    'Not yet certified': np.nan,
    'PhD Education': 'Doctor of Philosophy (PhD) in Education',
    'Phd': 'Doctor of Philosophy (PhD)',
    'Teacher Certificate': 'Certification of Education'
}

# Step 4: Fill in missing mappings using codes
for val in raw_qualifications:
    val_clean = val.strip()
    current_mapped_value = qualification_mapping.get(val_clean)

    if val_clean not in qualification_mapping or pd.isna(current_mapped_value):
        if val_clean in qual_code_to_name:
            qualification_mapping[val_clean] = qual_code_to_name[val_clean]

# Step 5: Auto-mapping via exact, case-insensitive, or fuzzy match
for val in raw_qualifications:
    val_clean = val.strip()
    val_lower = val_clean.lower()

    if val_clean in qualification_mapping:
        continue

    if val_clean in valid_qual_names:
        qualification_mapping[val_clean] = val_clean
    elif val_lower in valid_qual_names_lower:
        qualification_mapping[val_clean] = valid_qual_names_lower[val_lower]
    else:
        close = get_close_matches(val_clean, valid_qual_names, n=1, cutoff=0.8)
        if close:
            qualification_mapping[val_clean] = close[0]
        else:
            qualification_mapping[val_clean] = "__UNMAPPED__"

# Step 6: Apply mapping
df_staff_filtered['CleanedQualification'] = (
    df_staff_filtered['Highest Qualification']
    .apply(lambda val: qualification_mapping.get(str(val).strip()) if pd.notna(val) else np.nan)
)

# Step 7: Summary
print("📋 Final qualification mapping:")
for k, v in sorted(qualification_mapping.items()):
    print(f"'{k}': '{v}',")

# Count only valid non-null qualifications
valid_mask = df_staff_filtered['CleanedQualification'].isin(valid_qual_names)
valid_count = valid_mask.sum()
total_rows = df_staff_filtered.shape[0]
null_count = df_staff_filtered['CleanedQualification'].isna().sum()
invalid_count = total_rows - valid_count - null_count

print(f"\n✅ Auto-mapped valid qualifications: {valid_count}")
print(f"❌ Unmapped qualifications (invalid): {invalid_count}")
print(f"⬜ Missing qualifications (null): {null_count}")

# Step 8: Frequency breakdown
qualification_counts = df_staff_filtered['CleanedQualification'].value_counts(dropna=False)
print("\n📊 Frequency breakdown of 'CleanedQualification':")
print(qualification_counts)

# Step 9: Sample invalid rows (not in valid list and not null and not marked as __UNMAPPED__)
invalid_qual_df = df_staff_filtered[
    ~df_staff_filtered['CleanedQualification'].isin(valid_qual_names) &
    df_staff_filtered['CleanedQualification'].notna() &
    (df_staff_filtered['CleanedQualification'] == "__UNMAPPED__")
]

print("\n🚫 Sample rows with invalid or unmapped 'Highest Qualification' values:")
display(
    invalid_qual_df[[
        'School', 'First Name', 'Last Name', 'Highest Qualification', 'CleanedQualification'
    ]].head(100)
)


# %%
from difflib import get_close_matches
import numpy as np

# Step 1: Load valid certification names and code-to-name mapping from lookup, filtered for 'RMI Certification'
cert_lookup = [
    entry for entry in core_lookups['teacherQuals']
    if entry.get('G') == 'RMI Certification'
]

valid_cert_names = {entry['N'].strip() for entry in cert_lookup}
valid_cert_names_lower = {v.lower(): v for v in valid_cert_names}
cert_code_to_name = {entry['C'].strip(): entry['N'].strip() for entry in cert_lookup}

# Step 2: Extract unique raw certification values
raw_certifications = df_staff_filtered['Certification Level'].dropna().astype(str).str.strip().unique()

# Step 3: Manual mapping (optional starter, you can extend this later as needed)
certification_mapping = {
    'No Certificate': np.nan,
    'None': np.nan,
    'Not Certified': np.nan,
    'Not yet certified': np.nan,
    'Peofessional III': 'Professional Certificate III',
    'Prfofessional III': 'Professional Certificate III',
    'Profesional I': 'Professional Certificate I',
    'Professioan Certificate': 'Professional Certificate I',
    'Professioanl Certificate': 'Professional Certificate I',
    'Professional': 'Professional Certificate I',
    'Professional 1': 'Professional Certificate I',
    'Professional 2': 'Professional Certificate II',
    'Professional 3': 'Professional Certificate III',
    'Professional Certifate 1': 'Professional Certificate I',
    'Professional Certificate': 'Professional Certificate I',
    'Professional Certificate 1': 'Professional Certificate I',
    'Professional Certificate I': 'Professional Certificate I',
    'Professional Certificate II': 'Professional Certificate II',
    'Professional I': 'Professional Certificate I',
    'Professional II': 'Professional Certificate II',
    'Professional III': 'Professional Certificate III',
    'Proficient': 'Provisional Certificate',
    'Provisional': 'Provisional Certificate',
    'Provisional 1': 'Provisional Certificate',
    'Provisional 2': 'Provisional Certificate',
    'Provisional Certificate': 'Provisional Certificate',
    'Provisional I': 'Provisional Certificate',
    'Provisonal': 'Provisional Certificate',
    'Provissional': 'Provisional Certificate',
    'Standard Certificate': np.nan,
    'UNknown': np.nan,
    'Unknown': np.nan,
    'None': np.nan,
    'Not Certified': np.nan,
    'Not yet certified': np.nan,
    'No Certificate': np.nan
}

# Step 4: Fill in missing mappings by matching certification codes
for val in raw_certifications:
    val_clean = val.strip()
    current_mapped_value = certification_mapping.get(val_clean)

    if val_clean not in certification_mapping or pd.isna(current_mapped_value):
        if val_clean in cert_code_to_name:
            certification_mapping[val_clean] = cert_code_to_name[val_clean]

# Step 5: Add auto-mapping for remaining unmatched
for raw_value in raw_certifications:
    val = raw_value.strip()
    val_lower = val.lower()

    if val in certification_mapping:
        continue  # already mapped

    # 1. Exact match
    if val in valid_cert_names:
        certification_mapping[val] = val
        continue

    # 2. Case-insensitive match
    if val_lower in valid_cert_names_lower:
        certification_mapping[val] = valid_cert_names_lower[val_lower]
        continue

    # 3. Fuzzy match
    close = get_close_matches(val, valid_cert_names, n=1, cutoff=0.8)
    if close:
        certification_mapping[val] = close[0]
    else:
        certification_mapping[val] = "__UNMAPPED__"

# Step 6: Tag originally missing values
df_staff_filtered['__cert_missing'] = df_staff_filtered['Certification Level'].isna()

# Step 7: Apply mapping
def map_certification(val):
    if pd.isna(val):
        return np.nan
    stripped = str(val).strip()
    return certification_mapping.get(stripped, "__UNMAPPED__")

df_staff_filtered['CleanedEdCertification'] = df_staff_filtered['Certification Level'].apply(map_certification)

# Step 8: Summary
print("📋 Final certification mapping:")
for k, v in sorted(certification_mapping.items()):
    print(f"'{k}': '{v}',")

valid_mask = df_staff_filtered['CleanedEdCertification'].isin(valid_cert_names)
valid_count = valid_mask.sum()
null_count = df_staff_filtered['__cert_missing'].sum()
invalid_count = (df_staff_filtered['CleanedEdCertification'] == "__UNMAPPED__").sum()

print(f"\n✅ Auto-mapped valid certifications: {valid_count}")
print(f"❌ Unmapped certifications (invalid): {invalid_count}")
print(f"⬜ Missing certifications (null): {null_count}")

# Step 9: Frequency breakdown
cert_counts = df_staff_filtered['CleanedEdCertification'].value_counts(dropna=False)
print("\n📊 Frequency breakdown of 'CleanedEdCertification':")
print(cert_counts)

# Step 10: Sample invalid rows
invalid_cert_df = df_staff_filtered[
    df_staff_filtered['CleanedEdCertification'] == "__UNMAPPED__"
]

print("\n🚫 Sample rows with invalid or unmapped 'Certification Level' values:")
display(invalid_cert_df[['School', 'First Name', 'Last Name', 'Certification Level', 'CleanedEdCertification']].head(100))

# Optional: clean up marker value if needed
df_staff_filtered['CleanedEdCertification'] = df_staff_filtered['CleanedEdCertification'].replace("__UNMAPPED__", np.nan)

df_staff_filtered.drop(columns='__cert_missing', inplace=True)


# %%
from difflib import get_close_matches
import numpy as np
import re

# Step 1: Load valid subject names from lookup
subject_lookup = censusworkbook_lookups.get('subjects', [])
valid_subject_names = {entry['N'].strip() for entry in subject_lookup if entry.get('N')}
valid_subject_names_lower = {name.lower(): name for name in valid_subject_names}

# Step 2: Extract raw field of study values
raw_fields = df_staff_filtered['Field of Study'].dropna().astype(str).str.strip().unique()

# Step 3: Start with mapping directly from Field of Study
field_mapping = {}

for raw_val in raw_fields:
    val = raw_val.strip()
    val_lower = val.lower()

    if val in valid_subject_names:
        field_mapping[val] = val
    elif val_lower in valid_subject_names_lower:
        field_mapping[val] = valid_subject_names_lower[val_lower]
    else:
        close = get_close_matches(val, valid_subject_names, n=1, cutoff=0.8)
        field_mapping[val] = close[0] if close else "__UNMAPPED__"

# Step 4: Apply field_mapping
df_staff_filtered['CleanedFieldofStudy'] = (
    df_staff_filtered['Field of Study']
    .apply(lambda val: field_mapping.get(str(val).strip()) if pd.notna(val) else np.nan)
)

# Step 5: Attempt to fill missing fields using 'Highest Qualification'
def extract_possible_subject(text):
    if pd.isna(text):
        return None
    text = str(text)
    # Look for a subject-like phrase following 'in' or 'of'
    matches = re.findall(r'in ([A-Za-z&\.\- ]+)|of ([A-Za-z&\.\- ]+)', text, flags=re.IGNORECASE)
    possible_phrases = [m[0] or m[1] for m in matches]
    for phrase in possible_phrases:
        cleaned = phrase.strip().lower()
        if cleaned in valid_subject_names_lower:
            return valid_subject_names_lower[cleaned]
        close = get_close_matches(cleaned, valid_subject_names, n=1, cutoff=0.8)
        if close:
            return close[0]
    return None

# Only fill in missing field of study values
df_staff_filtered['CleanedFieldofStudy'] = df_staff_filtered.apply(
    lambda row: extract_possible_subject(row['Highest Qualification']) if pd.isna(row['CleanedFieldofStudy']) else row['CleanedFieldofStudy'],
    axis=1
)

# Step 6: Replace unmapped with NaN
df_staff_filtered['CleanedFieldofStudy'] = df_staff_filtered['CleanedFieldofStudy'].replace("__UNMAPPED__", np.nan)

# Step 7: Report
print("\n📊 Frequency breakdown of 'CleanedFieldofStudy':")
print(df_staff_filtered['CleanedFieldofStudy'].value_counts(dropna=False))

print("\n🚫 Sample unmapped or questionable 'Field of Study':")
display(df_staff_filtered[
    df_staff_filtered['CleanedFieldofStudy'].isna() &
    df_staff_filtered['Field of Study'].notna()
][['School', 'First Name', 'Last Name', 'Field of Study', 'Highest Qualification']].head(100))


# %%
from difflib import get_close_matches
import numpy as np

# Step 1: Define valid values
valid_teaching_staff = {"Teaching Staff", "Non Teaching Staff"}
valid_teaching_staff_lower = {v.lower(): v for v in valid_teaching_staff}

# Step 2: Extract unique raw values
raw_teaching_values = df_staff_filtered['Teaching Staff'].dropna().astype(str).str.strip().unique()

# Step 3: Manual mapping starter (can be extended)
teaching_staff_mapping = {
    'Admin': 'Teaching Staff',  # ?
    'Administration': 'Non Teaching Staff',  # ?
    'Administrator': 'Teaching Staff',  # ?
    'Head Teacher': 'Teaching Staff',
    'N/A': np.nan,
    'Non Teaching': 'Non Teaching Staff',
    'Non Teaching Staff': 'Non Teaching Staff',
    'Non teaching Staff': 'Non Teaching Staff',
    'Non-Teaching': 'Non Teaching Staff',
    'Non-Teaching Staff': 'Non Teaching Staff',
    'NonTeaching': 'Non Teaching Staff',
    'NonTeaching Staff': 'Non Teaching Staff',
    'None': np.nan,
    'Principal': 'Teaching Staff',  # ?
    'SPED': 'Teaching Staff',  # ?
    'SpEd Teacher': 'Teaching Staff',  # ?
    "T'eaching Staff": 'Teaching Staff',
    'Teacher': 'Teaching Staff',
    'Teaching': 'Teaching Staff',
    'Teaching Staff': 'Teaching Staff',
    'Unknown': np.nan
}

# Step 4: Auto-mapping for remaining values
for val in raw_teaching_values:
    val_clean = val.strip()
    val_lower = val_clean.lower()

    if val_clean in teaching_staff_mapping:
        continue

    if val_clean in valid_teaching_staff:
        teaching_staff_mapping[val_clean] = val_clean
        continue

    if val_lower in valid_teaching_staff_lower:
        teaching_staff_mapping[val_clean] = valid_teaching_staff_lower[val_lower]
        continue

    close = get_close_matches(val_clean, valid_teaching_staff, n=1, cutoff=0.85)
    if close:
        teaching_staff_mapping[val_clean] = close[0]
    else:
        teaching_staff_mapping[val_clean] = "__UNMAPPED__"

# Step 5: Tag missing values
df_staff_filtered['__teaching_missing'] = df_staff_filtered['Teaching Staff'].isna()

# Step 6: Apply mapping
def map_teaching(val):
    if pd.isna(val):
        return np.nan
    stripped = str(val).strip()
    return teaching_staff_mapping.get(stripped, "__UNMAPPED__")

df_staff_filtered['CleanedTeachingStaff'] = df_staff_filtered['Teaching Staff'].apply(map_teaching)

from difflib import get_close_matches

# Step 7: Fill in 'Teaching Staff' if CleanedTeachingStaff is still NaN and Job Title contains words similar to "teach"
def resembles_teaching(text):
    if pd.isna(text):
        return False
    words = str(text).lower().split()
    for word in words:
        if get_close_matches(word, ['teach'], n=1, cutoff=0.7):  # lowered cutoff for lenient matching
            return True
    return False

mask_infer_teaching = (
    df_staff_filtered['CleanedTeachingStaff'].isna() &
    df_staff_filtered['Job Title'].apply(resembles_teaching)
)

df_staff_filtered.loc[mask_infer_teaching, 'CleanedTeachingStaff'] = "Teaching Staff"


# Step 8: Summary
print("📋 Final teaching staff mapping:")
for k, v in sorted(teaching_staff_mapping.items()):
    print(f"'{k}': '{v}',")

valid_mask = df_staff_filtered['CleanedTeachingStaff'].isin(valid_teaching_staff)
valid_count = valid_mask.sum()
null_count = df_staff_filtered['__teaching_missing'].sum()
invalid_count = (df_staff_filtered['CleanedTeachingStaff'] == "__UNMAPPED__").sum()

print(f"\n✅ Auto-mapped valid Teaching Staff values: {valid_count}")
print(f"❌ Unmapped values (invalid): {invalid_count}")
print(f"⬜ Missing values (null): {null_count}")

# Step 9: Frequency breakdown
teaching_counts = df_staff_filtered['CleanedTeachingStaff'].value_counts(dropna=False)
print("\n📊 Frequency breakdown of 'CleanedTeachingStaff':")
print(teaching_counts)

# Step 10: Sample invalid and missing rows

# Invalid (unmapped) values
invalid_teaching_df = df_staff_filtered[
    df_staff_filtered['CleanedTeachingStaff'] == "__UNMAPPED__"
]

print("\n🚫 Sample rows with invalid or unmapped 'Teaching Staff' values:")
display(invalid_teaching_df[['School', 'First Name', 'Last Name', 'Teaching Staff', 'CleanedTeachingStaff']].head(100))

# Missing (NaN) values
missing_teaching_df = df_staff_filtered[
    df_staff_filtered['CleanedTeachingStaff'].isna()
]

print("\n⬜ Sample rows with missing 'Teaching Staff' values:")
display(missing_teaching_df[['School', 'First Name', 'Last Name', 'Teaching Staff', 'CleanedTeachingStaff']])

# Step 11: Final cleanup
df_staff_filtered['CleanedTeachingStaff'] = df_staff_filtered['CleanedTeachingStaff'].replace("__UNMAPPED__", np.nan)
df_staff_filtered.drop(columns='__teaching_missing', inplace=True)


# %%
from difflib import get_close_matches
import numpy as np

# Step 1: Load valid job titles from lookup
role_lookup = censusworkbook_lookups.get('censusTeacherRoles', [])
valid_roles = {entry['N'].strip() for entry in role_lookup if entry.get('N')}
valid_roles_lower = {role.lower(): role for role in valid_roles}

# Step 2: Extract unique raw job titles
raw_job_titles = df_staff_filtered['Job Title'].dropna().astype(str).str.strip().unique()

# Step 3: Manual mapping starter (can be extended later)
job_title_mapping = {
    '10th Counselor': 'Counselor',
    '12th Counselor': 'Counselor',
    '9th Counselor': 'Counselor',
    'Academy Reistrar': 'Registrar',
    'Accountant': 'Accountant',
    'Accounting Clerk': 'Accountant',
    'Acting Principal/Classroom Teacher': 'Classroom Teacher I',
    'Admin. Secretary': 'Secretary',
    'Administrative Assisstant': '__UNMAPPED__',
    'Adminstrator': '__UNMAPPED__',
    'Adminstrator/Classroom Teacher': 'Classroom Teacher I',
    'Assistant Clerk': '__UNMAPPED__',
    'Auto -teacher': 'Classroom Teacher I',
    'Boat operator': '__UNMAPPED__',
    'Budget Officer': '__UNMAPPED__',
    'Bus Driver': 'Bus Driver',
    'CLASSROOM TEACHER I': 'Classroom Teacher I',
    'CLASSROOM TEACHER II': 'Classroom Teacher II',
    'COOK': 'Cook',
    'Cafeteria Worker': '__UNMAPPED__',
    'Cafeteria Worker Supervisor': '__UNMAPPED__',
    'Chaplain': '__UNMAPPED__',
    'Chauffer/Driver': 'Bus Driver',
    'Clasroom Teacher': 'Classroom Teacher I',
    'Class Room Teacher/V-Principal': 'Classroom Teacher I',
    'Classroom Teacer': 'Classroom Teacher I',
    'Classroom Teache Aide': 'Classroom Teacher I',
    'Classroom Teacher': 'Classroom Teacher I',
    'Classroom Teacher Aide': 'Classroom Teacher I',
    'Classroom Teacher I': 'Classroom Teacher I',
    'Classroom Teacher/Head Teacher': 'Classroom Teacher I',
    'Classroom teacher': 'Classroom Teacher I',
    'Classroon Teacher': 'Classroom Teacher I',
    'Cook': 'Cook',
    'Counselor': 'Counselor',
    'Counselor / Classroom Teacher': 'Classroom Teacher I',
    'Counselor/Classroom Teacher': 'Classroom Teacher I',
    'Curriculum Specialist': '__UNMAPPED__',
    'Custidoan': '__UNMAPPED__',
    'Custodian': '__UNMAPPED__',
    'Data Specialist': 'Data Specialist',
    'Director': 'Director',
    'Dorm Cook': 'Cook',
    'Driver': 'Bud Driver',
    'Elective Teacher': 'Teacher Aide',
    'Female Counselor': 'Counselor',
    'Financial Director': '__UNMAPPED__',
    'Fiscal Officer': 'Fiscal Officer',
    'Gardener Teacher': 'Classroom Teacher I',
    'Girl''s Counselor': 'Counselor',
    'H-Teacher/Teacher': 'Head Teacher I',
    'Handy Man': '__UNMAPPED__',
    'Head Teacher': 'Head Teacher I',
    'Head Teacher/ Classroom Teacher': 'Head Teacher I',
    'Head Teacher/(Principal)Classroom Teacher': 'Head Teacher I',
    'Head Teacher/Clasroom Teacher': 'Head Teacher I',
    'Head Teacher/Classroom Teacher': 'Head Teacher I',
    'Head Teacher/Teacher': 'Head Teacher I',
    'Heard Teacher/Classroom Teacher': 'Head Teacher I',
    'House Baba': 'House Parent',
    'House mama': 'House Parent',
    'Housefather': 'House Parent',
    'Housemother': 'House Parent',
    'Houseparent': 'House Parent',
    'IT': 'Information Technology Officer',
    'IT/Computer/ Teacher': 'Information Technology Officer',
    'Information Technology Officer': 'Information Technology Officer',
    'It Trainee/STEM Specialist': 'Information Technology Officer',
    'Janitor': '__UNMAPPED__',
    'Janitor/ Support': '__UNMAPPED__',
    'Janitoress support': '__UNMAPPED__',
    'Janitress': '__UNMAPPED__',
    'Kinder V-Principal': '__UNMAPPED__',
    'Kitchen Staff': '__UNMAPPED__',
    'Librarian': 'Librarian',
    'MIHS Cook': 'Cook',
    'MIHS Security': 'Security Guard',
    'MLA Teacher': 'Classroom Teacher I',
    'MLA Techer': 'Classroom Teacher I',
    'Maintenace': 'Maintenance',
    'Maintenance': 'Maintenance',
    'Maintenance Support': 'Maintenance',
    'Maintenance Support/Driver': 'Maintenance',
    'Male Counselor': 'Counselor',
    'Mantenance supervisor': '__UNMAPPED__',
    'Mechanic': '__UNMAPPED__',
    'N/A': np.nan,
    'None': np.nan,
    'Not Assigned': np.nan,
    'Nurse': 'Nurse',
    'Nurse / CT': 'Nurse',
    'Office Assistant': '__UNMAPPED__',
    'P.E. Teacher': 'Classroom Teacher I',
    'Pre-9th (counselor)': 'Counselor',
    'Principal': 'Principal (Primary)',
    'Principal (Secondary)': 'Principal (Secondary)',
    'Principal/Classroom Teacher': 'Head Teacher I',
    'Principal/Clssroom Teacher': 'Head Teacher I',
    'Principal/Teacher': 'Principal (Primary)',
    'Prinicpal/Classroom Teacher': 'Head Teacher I',
    'Recreations/Sports/PE': '__UNMAPPED__',
    'Registrar': 'Registrar',
    'SEP': '__UNMAPPED__',
    'SPED': 'SEP Coordinator',
    'SPED Coordinator': 'SEP Coordinator',
    'SPED Teacher': 'Classroom Teacher I',
    'SPED.Teacher': 'Classroom Teacher I',
    'SPED/Classroom Teaher': 'Classroom Teacher I',
    'SPED/RSA': 'SEP Coordinator',
    'STEM Specialist': '__UNMAPPED__',
    'School Admin -Vice Principal': 'Vice Principal (Primary)',
    'School Admin- Principal': 'Principal (Primary)',
    'School Counselor': 'Counselor',
    'School Nurse': 'Nurse',
    'School Secretary': 'Secretary',
    'Secretary': 'Secretary',
    'Secretary/ Classroom Teacher': 'Classroom Teacher I',
    'Secretary/Collector': 'Secretary',
    'Security': 'Security Guard',
    'Security Guard': 'Security Guard',
    'Securtity': 'Security Guard',
    'SpEd Teacher': 'Classroom Teacher I',
    'Special Education Teacher': 'Classroom Teacher I',
    'Special Worker': 'Special Worker',
    'Sped Teacher': 'Classroom Teacher I',
    'Support Staff': '__UNMAPPED__',
    'Teacher': 'Classroom Teacher I',
    'Teacher Aid': 'Teacher Aide',
    'Teacher Aide': 'Teacher Aide',
    'Unassigned': np.nan,
    'Unknown': np.nan,
    'Unkown': np.nan,
    'V-Principal': 'Vice Principal (Primary)',
    'V.Principal': 'Vice Principal (Primary)',
    'Vice Pri': 'Vice Principal (Primary)',
    'Vice Principal': 'Vice Principal (Primary)',
    'Vice Principal (Primary)': 'Vice Principal (Primary)',
    'Vice Principal/Classroom Teacher': 'Head Teacher I',
    'Vice principal': 'Vice Principal (Primary)',
    'Vice-Principal': 'Vice Principal (Primary)',
    'WASC Coordinator': 'SEP Coordinator',
    'Watchman': '__UNMAPPED__',
    'None': np.nan,
    'N/A': np.nan,
    'Not Assigned': np.nan,
    'Unassigned': np.nan,
    'Unkown': np.nan,
    'Unknown': np.nan,
}

# Step 4: Fill in missing mappings by matching
for val in raw_job_titles:
    val_clean = val.strip()
    current_mapped_value = job_title_mapping.get(val_clean)

    if val_clean not in job_title_mapping or pd.isna(current_mapped_value):
        # No code-to-name mapping needed for roles, so we skip that step

        # Step 5: Auto-mapping
        val_lower = val_clean.lower()

        # Skip if already mapped
        if val_clean in job_title_mapping:
            continue

        # 1. Exact match
        if val_clean in valid_roles:
            job_title_mapping[val_clean] = val_clean
            continue

        # 2. Case-insensitive match
        if val_lower in valid_roles_lower:
            job_title_mapping[val_clean] = valid_roles_lower[val_lower]
            continue

        # 3. Fuzzy match
        close = get_close_matches(val_clean, valid_roles, n=1, cutoff=0.8)
        if close:
            job_title_mapping[val_clean] = close[0]
        else:
            job_title_mapping[val_clean] = "__UNMAPPED__"

# Step 6: Tag missing values
df_staff_filtered['__job_missing'] = df_staff_filtered['Job Title'].isna()

# Step 7: Apply mapping
def map_job_title(val):
    if pd.isna(val):
        return np.nan
    stripped = str(val).strip()
    return job_title_mapping.get(stripped, "__UNMAPPED__")

df_staff_filtered['CleanedJobTitle'] = df_staff_filtered['Job Title'].apply(map_job_title)

# Step 8: Summary
print("📋 Final job title mapping:")
for k, v in sorted(job_title_mapping.items()):
    print(f"'{k}': '{v}',")

valid_mask = df_staff_filtered['CleanedJobTitle'].isin(valid_roles)
valid_count = valid_mask.sum()
null_count = df_staff_filtered['__job_missing'].sum()
invalid_count = (df_staff_filtered['CleanedJobTitle'] == "__UNMAPPED__").sum()

print(f"\n✅ Auto-mapped valid job titles: {valid_count}")
print(f"❌ Unmapped job titles (invalid): {invalid_count}")
print(f"⬜ Missing job titles (null): {null_count}")

# Step 9: Frequency breakdown
job_counts = df_staff_filtered['CleanedJobTitle'].value_counts(dropna=False)
print("\n📊 Frequency breakdown of 'CleanedJobTitle':")
print(job_counts)

# Step 10: Sample invalid rows (unmapped)
invalid_job_df = df_staff_filtered[
    df_staff_filtered['CleanedJobTitle'] == "__UNMAPPED__"
]

print("\n🚫 Sample rows with invalid or unmapped 'Job Title' values:")
display(invalid_job_df[['School', 'First Name', 'Last Name', 'Job Title', 'CleanedJobTitle', 'CleanedTeachingStaff']])

# Step 10.1: Sample rows with NaN
nan_job_df = df_staff_filtered[
    df_staff_filtered['CleanedTeachingStaff'].isna()
]

print("\n⬜ Sample rows with missing 'CleanedTeachingStaff' values (NaN):")
display(nan_job_df[['School', 'First Name', 'Last Name', 'Job Title', 'CleanedJobTitle', 'CleanedTeachingStaff']])

# Optional: Replace marker value
df_staff_filtered['CleanedJobTitle'] = df_staff_filtered['CleanedJobTitle'].replace("__UNMAPPED__", np.nan)

# Clean up temp column
df_staff_filtered.drop(columns='__job_missing', inplace=True)


# %%
from difflib import get_close_matches
import numpy as np

# Step 1: Define valid values
valid_status_values = {"Active", "Inactive"}
valid_status_lower = {v.lower(): v for v in valid_status_values}

# Step 2: Extract unique raw values
raw_status_values = df_staff_filtered['Employment Status'].dropna().astype(str).str.strip().unique()

# Step 3: Manual mapping starter (extend as needed)
status_mapping = {
    'Active': 'Active',
    'Inactive': 'Inactive',
    'Actve': 'Active',
    'InActive': 'Inactive',
    'Not Active': 'Inactive',
    'Retired': 'Inactive',
    'Still teaching': 'Active',
    'Unknown': np.nan,
    'None': np.nan,
    'N/A': np.nan,
    '': np.nan,
}

# Step 4: Auto-mapping for remaining values
for val in raw_status_values:
    val_clean = val.strip()
    val_lower = val_clean.lower()

    if val_clean in status_mapping:
        continue

    # 1. Exact match
    if val_clean in valid_status_values:
        status_mapping[val_clean] = val_clean
        continue

    # 2. Case-insensitive match
    if val_lower in valid_status_lower:
        status_mapping[val_clean] = valid_status_lower[val_lower]
        continue

    # 3. Fuzzy match
    close = get_close_matches(val_clean, valid_status_values, n=1, cutoff=0.8)
    if close:
        status_mapping[val_clean] = close[0]
    else:
        status_mapping[val_clean] = "__UNMAPPED__"

# Step 5: Tag missing values
df_staff_filtered['__status_missing'] = df_staff_filtered['Employment Status'].isna()

# Step 6: Apply mapping
def map_status(val):
    if pd.isna(val):
        return np.nan
    stripped = str(val).strip()
    return status_mapping.get(stripped, "__UNMAPPED__")

df_staff_filtered['CleanedEmploymentStatus'] = df_staff_filtered['Employment Status'].apply(map_status)

# Step 7: Summary
print("📋 Final employment status mapping:")
for k, v in sorted(status_mapping.items()):
    print(f"'{k}': '{v}',")

valid_mask = df_staff_filtered['CleanedEmploymentStatus'].isin(valid_status_values)
valid_count = valid_mask.sum()
null_count = df_staff_filtered['__status_missing'].sum()
invalid_count = (df_staff_filtered['CleanedEmploymentStatus'] == "__UNMAPPED__").sum()

print(f"\n✅ Auto-mapped valid statuses: {valid_count}")
print(f"❌ Unmapped statuses (invalid): {invalid_count}")
print(f"⬜ Missing statuses (null): {null_count}")

# Step 8: Frequency breakdown
status_counts = df_staff_filtered['CleanedEmploymentStatus'].value_counts(dropna=False)
print("\n📊 Frequency breakdown of 'CleanedEmploymentStatus':")
print(status_counts)

# Step 9: Sample invalid and missing rows
invalid_status_df = df_staff_filtered[
    df_staff_filtered['CleanedEmploymentStatus'] == "__UNMAPPED__"
]
print("\n🚫 Sample rows with invalid or unmapped 'Employment Status' values:")
display(invalid_status_df[['School', 'First Name', 'Last Name', 'Employment Status', 'CleanedEmploymentStatus']].head(100))

missing_status_df = df_staff_filtered[
    df_staff_filtered['CleanedEmploymentStatus'].isna()
]
print("\n⬜ Sample rows with missing 'Employment Status' values:")
display(missing_status_df[['School', 'First Name', 'Last Name', 'Employment Status', 'CleanedEmploymentStatus']].head(100))

# Step 10: Final cleanup
df_staff_filtered['CleanedEmploymentStatus'] = df_staff_filtered['CleanedEmploymentStatus'].replace("__UNMAPPED__", np.nan)
df_staff_filtered.drop(columns='__status_missing', inplace=True)


# %%
from difflib import get_close_matches
import numpy as np

# Step 1: Extract valid organization names from lookup
org_lookup = censusworkbook_lookups['organizations']
valid_org_names = {entry['N'].strip() for entry in org_lookup}
valid_org_names_lower = {v.lower(): v for v in valid_org_names}

# Step 2: Extract unique raw values
raw_org_values = df_staff_filtered['Organization'].dropna().astype(str).str.strip().unique()

# Step 3: Manual mapping starter (extend as needed)
org_mapping = {
    '': np.nan,
    'MOE': 'Ministry of Education',
    'Ministry Education': 'Ministry of Education',
    'N/A': np.nan,
    'None': np.nan,
    'PSS': 'PSS',
    'Private Schools': 'Private School Teacher',
    'Unknown': np.nan,
}

# Step 4: Auto-mapping for remaining values
for val in raw_org_values:
    val_clean = val.strip()
    val_lower = val_clean.lower()

    if val_clean in org_mapping:
        continue

    # 1. Exact match
    if val_clean in valid_org_names:
        org_mapping[val_clean] = val_clean
        continue

    # 2. Case-insensitive match
    if val_lower in valid_org_names_lower:
        org_mapping[val_clean] = valid_org_names_lower[val_lower]
        continue

    # 3. Fuzzy match
    close = get_close_matches(val_clean, valid_org_names, n=1, cutoff=0.85)
    if close:
        org_mapping[val_clean] = close[0]
    else:
        org_mapping[val_clean] = "__UNMAPPED__"

# Step 5: Tag missing values
df_staff_filtered['__org_missing'] = df_staff_filtered['Organization'].isna()

# Step 6: Apply mapping
def map_organization(val):
    if pd.isna(val):
        return np.nan
    stripped = str(val).strip()
    return org_mapping.get(stripped, "__UNMAPPED__")

df_staff_filtered['CleanedOrganization'] = df_staff_filtered['Organization'].apply(map_organization)

# Step 7: Infer organization from CleanedSchool via all_schools
school_auth_map = {
    school['schName'].strip(): school['schAuth'].strip() if school.get('schAuth') else None
    for school in all_schools
}

# Function to infer from schAuth
def infer_org_from_school(school_name):
    auth = school_auth_map.get(str(school_name).strip())
    if auth == "PSS":
        return "PSS"
    elif auth:
        return "Private School Teacher"
    return np.nan

# Mask: only rows that still need inference
mask_needs_inference = df_staff_filtered['CleanedOrganization'].isna() & df_staff_filtered['CleanedSchool'].notna()

# Apply inference
inferred_orgs = df_staff_filtered.loc[mask_needs_inference, 'CleanedSchool'].apply(infer_org_from_school)

# Assign back to the DataFrame
df_staff_filtered.loc[mask_needs_inference, 'CleanedOrganization'] = inferred_orgs

# Count results
pss_count = (inferred_orgs == "PSS").sum()
private_count = (inferred_orgs == "Private School Teacher").sum()
null_count = inferred_orgs.isna().sum()

# Identify distinct schools that led to "Private School Teacher"
private_school_names = df_staff_filtered.loc[
    mask_needs_inference & (inferred_orgs == "Private School Teacher"),
    'CleanedSchool'
].dropna().unique()

# Summary printout
print("\n🏫 Inferred 'Organization' from 'CleanedSchool':")
print(f"   ✅ PSS: {pss_count}")
print(f"   🏫 Private School Teacher: {private_count}")
print(f"       From schools: {', '.join(sorted(private_school_names))}")
print(f"   ⬜ Still unknown (no match or no auth info): {null_count}")


# Step 8: Summary
print("📋 Final organization mapping:")
for k, v in sorted(org_mapping.items()):
    print(f"'{k}': '{v}',")

valid_mask = df_staff_filtered['CleanedOrganization'].isin(valid_org_names)
valid_count = valid_mask.sum()
null_count = df_staff_filtered['__org_missing'].sum()
invalid_count = (df_staff_filtered['CleanedOrganization'] == "__UNMAPPED__").sum()

print(f"\n✅ Auto-mapped valid organizations: {valid_count}")
print(f"❌ Unmapped values (invalid): {invalid_count}")
print(f"⬜ Missing values (null): {null_count}")

# Step 9: Frequency breakdown
org_counts = df_staff_filtered['CleanedOrganization'].value_counts(dropna=False)
print("\n📊 Frequency breakdown of 'CleanedOrganization':")
print(org_counts)

# Step 10: Sample invalid and missing rows
invalid_org_df = df_staff_filtered[
    df_staff_filtered['CleanedOrganization'] == "__UNMAPPED__"
]
print("\n🚫 Sample rows with invalid or unmapped 'Organization' values:")
display(invalid_org_df[['School', 'First Name', 'Last Name', 'Organization', 'CleanedOrganization']].head(100))

missing_org_df = df_staff_filtered[
    df_staff_filtered['CleanedOrganization'].isna()
]
print("\n⬜ Sample rows with missing 'Organization' values:")
display(missing_org_df[['School', 'First Name', 'Last Name', 'Organization', 'CleanedOrganization']].head(100))

# Step 11: Final cleanup
df_staff_filtered['CleanedOrganization'] = df_staff_filtered['CleanedOrganization'].replace("__UNMAPPED__", np.nan)
df_staff_filtered.drop(columns='__org_missing', inplace=True)


# %%
import numpy as np

# Step 1: Define the columns to clean
grade_columns = [
    "ECE", "Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5",
    "Grade 6", "Grade 7", "Grade 8", "Pre-9", "Grade 9", "Grade 10",
    "Grade 11", "Grade 12", "Admin", "Other"
]

print("📋 Cleaning teaching assignment markers (interpreting any non-blank as 'x'):")

cleaned_column_names = []

# Step 2: Clean each column
for col in grade_columns:
    cleaned_col = f"Cleaned{col.replace(' ', '').replace('-', '')}"
    cleaned_column_names.append(cleaned_col)

    # Step 3: Convert original values to lowercase and strip spaces
    original = df_staff_filtered[col].astype(str).str.strip().str.lower()

    # Step 4: Identify raw value breakdown before cleaning
    counts = original.value_counts(dropna=False)
    print(f"\n📊 Column: '{col}' — raw value breakdown:")
    print(counts)

    # Step 5: Treat any non-empty string (even weird ones like 'yes', 's', etc.) as 'x'
    df_staff_filtered[cleaned_col] = original.apply(lambda x: "x" if x and x != "nan" else np.nan)

    # Step 6: Summary stats
    num_x = (df_staff_filtered[cleaned_col] == 'x').sum()
    num_na = df_staff_filtered[cleaned_col].isna().sum()
    total = len(df_staff_filtered)

    print(f"✅ Cleaned column: '{cleaned_col}'")
    print(f"   ✔ Interpreted 'x' values: {num_x}")
    print(f"   ⬜ Blank or empty (set to NaN): {num_na}")
    print(f"   📦 Total: {total}")

print("\n✅ All cleaned columns created:")
print(cleaned_column_names)


# %%
# Ensure CleanedOther column exists and is of type object
if 'CleanedOther' not in df_staff_filtered.columns:
    df_staff_filtered['CleanedOther'] = np.nan

df_staff_filtered['CleanedOther'] = df_staff_filtered['CleanedOther'].astype(object)

# Apply inference safely
mask_infer_other = (
    (df_staff_filtered['CleanedTeachingStaff'] == "Non Teaching Staff") &
    (df_staff_filtered['CleanedOther'].isna())
)

df_staff_filtered.loc[mask_infer_other, 'CleanedOther'] = "x"

# Summary of changes
inferred_other_count = mask_infer_other.sum()
total_other_x = (df_staff_filtered['CleanedOther'] == "x").sum()
total_rows = len(df_staff_filtered)

print("\n🧩 Inference for 'CleanedOther':")
print(f"   🆕 Inferred from 'Non Teaching Staff': {inferred_other_count}")
print(f"   ✅ Total 'x' values in CleanedOther: {total_other_x}")
print(f"   📦 Total records: {total_rows}")


# %%
from difflib import get_close_matches

# Step 1: Extract valid values from the lookup
teacher_type_entries = censusworkbook_lookups['teacherRegStatus']
valid_teacher_types = [entry['N'] for entry in teacher_type_entries]
valid_teacher_types_set = set(valid_teacher_types)
valid_teacher_types_lower = {v.lower(): v for v in valid_teacher_types}

# Step 2: Build mapping from raw values
teacher_type_mapping = {
    'Regular': 'Regular',
    'Special Ed': 'Special Ed',
    'Teacher-Aide': 'Assistant'
}

raw_values = df_staff_filtered['Teacher-Type'].dropna().unique()

for val in raw_values:
    val_clean = str(val).strip()
    val_lower = val_clean.lower()

    # 1. Exact match
    if val_clean in valid_teacher_types_set:
        teacher_type_mapping[val_clean] = val_clean
        continue

    # 2. Case-insensitive match
    if val_lower in valid_teacher_types_lower:
        teacher_type_mapping[val_clean] = valid_teacher_types_lower[val_lower]
        continue

    # 3. Fuzzy match
    close = get_close_matches(val_clean, valid_teacher_types, n=1, cutoff=0.85)
    if close:
        teacher_type_mapping[val_clean] = close[0]
    else:
        teacher_type_mapping[val_clean] = "__UNMAPPED__"

# Step 3: Tag missing values
df_staff_filtered['__teacher_type_missing'] = df_staff_filtered['Teacher-Type'].isna()

# Step 4: Apply mapping
def map_teacher_type(val):
    if pd.isna(val):
        return np.nan
    val_stripped = str(val).strip()
    return teacher_type_mapping.get(val_stripped, "__UNMAPPED__")

df_staff_filtered['CleanedTeacherType'] = df_staff_filtered['Teacher-Type'].apply(map_teacher_type)

# Step 5: Summary
valid_count = df_staff_filtered['CleanedTeacherType'].isin(valid_teacher_types).sum()
null_count = df_staff_filtered['__teacher_type_missing'].sum()
invalid_count = (df_staff_filtered['CleanedTeacherType'] == "__UNMAPPED__").sum()

print("\n🧑‍🏫 Cleaned 'Teacher-Type' → 'CleanedTeacherType':")
print(f"   ✅ Valid entries: {valid_count}")
print(f"   ❌ Unmapped entries: {invalid_count}")
print(f"   ⬜ Missing entries: {null_count}")

# Step 6: Mapping summary
print("\n📋 Mapping used:")
for k, v in sorted(teacher_type_mapping.items()):
    print(f"'{k}': '{v}'")

# Step: Infer 'CleanedTeacherType' as 'Regular' if missing but marked as Teaching Staff
mask_infer_regular = (
    df_staff_filtered['CleanedTeacherType'].isna() &
    (df_staff_filtered['CleanedTeachingStaff'] == "Teaching Staff")
)

df_staff_filtered.loc[mask_infer_regular, 'CleanedTeacherType'] = "Regular"

# Summary
print(f"✅ Inferred 'Regular' for missing 'CleanedTeacherType' where staff is marked as 'Teaching Staff': {mask_infer_regular.sum()} rows updated.")

# Step 7: Frequency breakdown
print("\n📊 Frequency breakdown:")
print(df_staff_filtered['CleanedTeacherType'].value_counts(dropna=False))

# Step 8: Invalid samples
invalid_df = df_staff_filtered[df_staff_filtered['CleanedTeacherType'] == "__UNMAPPED__"]
print("\n🚫 Sample rows with unmapped 'Teacher-Type' values:")
display(invalid_df[['School', 'First Name', 'Last Name', 'Teacher-Type', 'CleanedTeacherType']].head(100))

# Step 9: Missing samples
missing_df = df_staff_filtered[df_staff_filtered['CleanedTeacherType'].isna()]
print("\n⬜ Sample rows with missing 'Teacher-Type' values:")
display(missing_df[['School', 'First Name', 'Last Name', 'Teacher-Type', 'CleanedTeacherType']].head(100))

# Step 10: Final cleanup
df_staff_filtered['CleanedTeacherType'] = df_staff_filtered['CleanedTeacherType'].replace("__UNMAPPED__", np.nan)
df_staff_filtered.drop(columns="__teacher_type_missing", inplace=True)


# %%
# Step 1: Copy and preprocess the source column
raw_dates = df_staff_filtered['Date of Hire'].astype(str).str.strip()

# Step 2: Try parsing full dates directly
cleaned_dates = pd.to_datetime(raw_dates, errors='coerce', dayfirst=False)

# Step 3: Handle year-only values like "2020"
year_only_mask = cleaned_dates.isna() & raw_dates.str.fullmatch(r"\d{4}")
parsed_years = pd.to_datetime(raw_dates[year_only_mask] + "-10-01", errors='coerce')
cleaned_dates[year_only_mask] = parsed_years

# Step 4: Fill in missing values from df_teacher_recent_survey_data using name + schNo match
missing_mask = cleaned_dates.isna()

df_recent = df_teacher_recent_survey_data[['tGiven', 'tSurname', 'tPayroll', 'tDatePSAppointed', 'schNo']].copy()

df_recent['tKey'] = (
    df_recent['tGiven'].astype(str).str.strip().str.lower() + "|" +
    df_recent['tSurname'].astype(str).str.strip().str.lower() + "|" +
    df_recent['schNo'].astype(str).str.strip().str.upper()
)

df_staff_filtered['nameKey'] = (
    df_staff_filtered['First Name'].astype(str).str.strip().str.lower() + "|" +
    df_staff_filtered['Last Name'].astype(str).str.strip().str.lower() + "|" +
    df_staff_filtered['schNo'].astype(str).str.strip().str.upper()
)

# Check for duplicates in tKey
duplicate_keys = df_recent[df_recent.duplicated(subset='tKey', keep=False)]

print("\n🚨 Duplicate tKey entries found in df_teacher_recent_survey_data:")
print(duplicate_keys.sort_values('tKey').to_string(index=False))

# Prioritize rows with non-null tDatePSAppointed and tPayroll
df_recent = df_recent.sort_values(
    by=['tDatePSAppointed', 'tPayroll'], 
    ascending=[False, False]
)

# Then drop duplicates, keeping the first (which now has more complete data)
df_recent = df_recent.drop_duplicates(subset='tKey', keep='first')

matched_dates = df_staff_filtered.loc[missing_mask, 'nameKey'].map(
    df_recent.set_index('tKey')['tDatePSAppointed']
)

matched_dates_parsed = pd.to_datetime(matched_dates, errors='coerce')
cleaned_dates[missing_mask] = matched_dates_parsed

# Assign final cleaned date column
df_staff_filtered['CleanedDateofHire'] = cleaned_dates

# Summary
num_parsed = cleaned_dates.notna().sum()
num_total = len(cleaned_dates)
num_filled_from_recent = matched_dates_parsed.notna().sum()

print("📅 Date of Hire Cleaning Summary:")
print(f"✅ Total parsed or inferred: {num_parsed} / {num_total}")
print(f"   🔄 Filled from recent survey data (name + school match): {num_filled_from_recent}")
print(f"   ⬜ Still missing: {(df_staff_filtered['CleanedDateofHire'].isna()).sum()}")


# %%
# Step 1: Create unique keys for matching
df_recent = df_teacher_recent_survey_data[['tGiven', 'tSurname', 'tPayroll', 'tchSalary', 'schNo']].copy()

df_recent['tKey'] = (
    df_recent['tGiven'].astype(str).str.strip().str.lower() + "|" +
    df_recent['tSurname'].astype(str).str.strip().str.lower() + "|" +
    df_recent['schNo'].astype(str).str.strip().str.upper()
)

df_staff_filtered['nameKey'] = (
    df_staff_filtered['First Name'].astype(str).str.strip().str.lower() + "|" +
    df_staff_filtered['Last Name'].astype(str).str.strip().str.lower() + "|" +
    df_staff_filtered['schNo'].astype(str).str.strip().str.upper()
)

# Step 2: Check for duplicates in tKey
duplicate_keys_salary = df_recent[df_recent.duplicated(subset='tKey', keep=False)]
print("\n🚨 Duplicate tKey entries found in df_teacher_recent_survey_data (Salary context):")
print(duplicate_keys_salary.sort_values('tKey').to_string(index=False))

# Step 3: Prioritize rows with non-null tchSalary and drop duplicates
df_recent = df_recent.sort_values(
    by=['tchSalary', 'tPayroll'], 
    ascending=[False, False]
)

df_recent = df_recent.drop_duplicates(subset='tKey', keep='first')

# Step 4: Perform the mapping for missing salary values
salary_matched = df_staff_filtered['nameKey'].map(
    df_recent.set_index('tKey')['tchSalary']
)

# Step 5: Assign final cleaned salary
df_staff_filtered['CleanedAnnualSalary'] = pd.to_numeric(salary_matched, errors='coerce')

# Step 6: Summary
total_filled = df_staff_filtered['CleanedAnnualSalary'].notna().sum()
total_rows = len(df_staff_filtered)

print("💰 Annual Salary Cleaning Summary:")
print(f"✅ Total filled from matched teacher survey data: {total_filled} / {total_rows}")
print(f"⬜ Still missing: {(df_staff_filtered['CleanedAnnualSalary'].isna()).sum()}")


# %%
import xlwings as xw
import os
import shutil
import pandas as pd

# Paths
new_workbook_path = os.path.join(output_directory, empty_census_workbook_filename)
clean_workbook_path = os.path.join(output_directory, clean_census_workbook_filename)

# 🧹 Delete and copy clean workbook template (controlled by flag)
if delete_census_workbook_filename:
    if os.path.exists(new_workbook_path):
        try:
            os.remove(new_workbook_path)
            print(f"🗑️ Removed previous workbook: {new_workbook_path}")
        except Exception as e:
            raise RuntimeError(f"Failed to delete existing workbook: {new_workbook_path}\n{e}")

    try:
        shutil.copyfile(clean_workbook_path, new_workbook_path)
        print(f"📄 Copied clean template to: {new_workbook_path}")
    except Exception as e:
        raise RuntimeError(f"Failed to copy clean workbook template.\n{e}")
else:
    print("⚠️ Skipping deletion and copy of census workbook (using existing workbook).")


# 📌 Mapping cleaned columns → Excel columns
column_mapping = {
    'CleanedYear': 'SchoolYear',
    #'dName': 'Atoll / Island',
    'CleanedSchool': 'School Name',
    #'schNo': 'PSS School ID',
    'CleanedOrganization': 'Organization',
    #'Office': 'Office',
    'First Name': 'First Name',
    'Middle Name': 'Middle Name',
    'Last Name': 'Last Name',
    #'Full Name': 'Full Name',
    'Gender': 'Gender',
    'CleanedDateofBirth': 'Date of Birth',
    #'Age': 'Age',
    'CleanedCitizenship': 'Citizenship',
    'CleanedEthnicity': 'Ethnicity',
    'RMI SS#': 'RMI SSN',
    #'Other SS#': 'Other SSN',
    'CleanedQualification': 'Highest Qualification',
    'CleanedFieldofStudy': 'Field of Study',
    'Year of Completion': 'Year of Completion',
    'CleanedEdCertification': 'Highest Ed Certification',
    #'Year Of Completion.1': 'Year Of Completion2',
    'CleanedEmploymentStatus': 'Employment Status',
    'Reason': 'Reason',
    'CleanedJobTitle': 'Job Title',
    'CleanedOrganization': 'Organization',
    'CleanedTeachingStaff': 'Staff Type',
    'CleanedTeacherType': 'Teacher Type',
    'CleanedDateofHire': 'Date of Hire',
    'Date of Exit': 'Date Of Exit',
    'CleanedAnnualSalary': 'Annual Salary',
    'Funding Source': 'Funding Source',
    'CleanedECE': 'ECE',
    'CleanedGrade1': 'Grade 1',
    'CleanedGrade2': 'Grade 2',
    'CleanedGrade3': 'Grade 3',
    'CleanedGrade4': 'Grade 4',
    'CleanedGrade5': 'Grade 5',
    'CleanedGrade6': 'Grade 6',
    'CleanedGrade7': 'Grade 7',
    'CleanedGrade8': 'Grade 8',
    #'CleanedPre9': 'Grade 9',
    'CleanedGrade9': 'Grade 9',
    'CleanedGrade10': 'Grade 10',
    'CleanedGrade11': 'Grade 11',
    'CleanedGrade12': 'Grade 12',
    'CleanedAdmin': 'Admin',
    'CleanedOther': 'Other',
    #'Total Days Absence': 'Total Days Absence',
    #'Maths': 'Maths',
    #'Science': 'Science',
    #'Language': 'Language',
    #'Competency': 'Competency',
    #'Teach Mathematics': 'Teach Mathematics',
    #'Teach Language Arts': 'Teach Language Arts',
    #'Teach Social Studies': 'Teach Social Studies',
    #'Teach Sciences': 'Teach Sciences'
}

try:
    app = xw.App(visible=False)
    wb = app.books.open(new_workbook_path)
    
    # Get the correct worksheet
    sheet_name = 'SchoolStaff'
    ws = wb.sheets[sheet_name]
    
    # ✅ Unprotect the correct sheet
    ws.api.Unprotect()  # Add password if needed

    # Prepare data
    df_to_insert = df_staff_filtered[list(column_mapping.keys())].copy()
    df_to_insert.rename(columns=column_mapping, inplace=True)
    df_to_insert = df_to_insert.astype(object).where(pd.notna(df_to_insert), None)

    print('df_to_insert: ')
    display(df_to_insert.head(3))    
    print(df_to_insert.columns)
    
    # Locate header row
    header_row = 3
    excel_headers = ws.range((header_row, 1)).expand('right').value
    header_indices = {
        header: idx + 1 for idx, header in enumerate(excel_headers) if header in df_to_insert.columns
    }

    # Write to Excel
    start_row = header_row + 1
    num_rows = len(df_to_insert)
    invalid_columns = []

    for col_name, col_idx in header_indices.items():
        try:
            col_values = df_to_insert[col_name].tolist()
            ws.range((start_row, col_idx), (start_row + num_rows - 1, col_idx)).value = [[v] for v in col_values]
        except Exception as e:
            print(f"❌ Error in column: {col_name} (Excel column {col_idx})")
            print(e)
            invalid_columns.append(col_name)

    if len(invalid_columns) > 0:
        print(f"\n⚠️ Columns that failed to write: {invalid_columns}")

    wb.save()
finally:
    wb.close()
    app.quit()

print("✅ Staff data successfully injected into SchoolStaff sheet.")


# %%
df_staff_filtered.columns
