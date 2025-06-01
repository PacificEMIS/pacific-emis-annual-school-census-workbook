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

def load_config(config_path="config.json"):
    """Load configuration from a JSON file."""
    with open(config_path, 'r') as file:
        config = json.load(file)
    return config['output_directory'], config['source_workbook_filename'], config['census_workbook_filename']
    

# Test loading configuration
output_directory, source_workbook_filename, census_workbook_filename = load_config()
print("Configuration loaded successfully.")

# %%
# Load workbook

# Combine the directory and filename
workbook_path = os.path.join(output_directory, census_workbook_filename)

# Load the workbook (but not yet any sheet into memory)
xls = pd.ExcelFile(workbook_path)

# See the available sheet names
print("Sheets found:", xls.sheet_names)

# Load the "Students" sheet
df_students = pd.read_excel(workbook_path, sheet_name='Students', skiprows=2)
print("Student columns found:", df_students.columns)

# %%
df_students

# %%
import matplotlib.pyplot as plt

# %store -r core_lookups student_lookups censusworkbook_lookups

# Step 1: Extract census year info for 2025
census_info = next(x for x in censusworkbook_lookups['censusYears'] if x['svyYear'] == 2025)
census_date = pd.to_datetime(census_info['svyCensusDate'])
ps_age = census_info['svyPSAge']  # Age for Grade 1

# Step 2: Prepare ExpectedAge (aka Official Age) from grade level names
official_age_map = {
    entry['N']: ps_age - 1 + entry['YoEd']
    for entry in core_lookups['levels']
    if 'YoEd' in entry and pd.notna(entry['YoEd'])
}

df_students['ExpectedAge'] = df_students['Grade Level'].map(official_age_map)

# Step 3: Categorize AgeGroup based on reported Age vs. ExpectedAge
df_students['Age'] = pd.to_numeric(df_students['Age'], errors='coerce')

df_students['AgeGroup'] = df_students.apply(
    lambda row: (
        'Official Age' if pd.notna(row['Age']) and pd.notna(row['ExpectedAge']) and row['Age'] == row['ExpectedAge']
        else 'Under Age' if pd.notna(row['Age']) and pd.notna(row['ExpectedAge']) and row['Age'] < row['ExpectedAge']
        else 'Over Age' if pd.notna(row['Age']) and pd.notna(row['ExpectedAge']) and row['Age'] > row['ExpectedAge']
        else 'Unknown'
    ),
    axis=1
)

# Step 4: Group and plot
age_group_dist = df_students[
    df_students['Grade Level'].notna() & df_students['AgeGroup'].notna()
].groupby(['Grade Level', 'AgeGroup']).size().unstack(fill_value=0)

# Reorder grades if necessary
grade_order_clean = [entry['N'] for entry in core_lookups['levels'] if entry['N'] in age_group_dist.index]
age_group_dist = age_group_dist.reindex(grade_order_clean)

# Reorder columns to: Under Age, Official Age, Over Age
desired_order = ['Under Age', 'Official Age', 'Over Age']
available_order = [col for col in desired_order if col in age_group_dist.columns]
age_group_dist = age_group_dist[available_order]

# Plot 1: All age groups
ax1 = age_group_dist.plot(
    kind='bar', stacked=False, figsize=(12, 6),
    title='(Official Cleaned Data) Enrollment by Grade and Age Group',
    ylabel='Number of Students', xlabel='Grade Level',
    rot=45, grid=True, legend=True
)
for container in ax1.containers:
    ax1.bar_label(container, label_type='edge', fontsize=8)
plt.tight_layout()
plt.show()

# Plot 2: Only Under and Over Age
age_group_filtered = age_group_dist.drop(columns='Official Age', errors='ignore')
ax2 = age_group_filtered.plot(
    kind='bar', stacked=False, figsize=(12, 6),
    title='(Official Cleaned Data) Enrollment by Grade: Under & Over Age Only',
    ylabel='Number of Students', xlabel='Grade Level',
    rot=45, grid=True, legend=True
)
for container in ax2.containers:
    ax2.bar_label(container, label_type='edge', fontsize=8)
plt.tight_layout()
plt.show()


# %%
import matplotlib.pyplot as plt

# %store -r core_lookups student_lookups censusworkbook_lookups

# Step 1: Extract census year info for 2025
census_info = next(x for x in censusworkbook_lookups['censusYears'] if x['svyYear'] == 2025)
census_date = pd.to_datetime(census_info['svyCensusDate'])
ps_age = census_info['svyPSAge']  # Age for Grade 1

# Step 2: Prepare ExpectedAge (aka Official Age) from grade level names
official_age_map = {
    entry['N']: ps_age - 1 + entry['YoEd']
    for entry in core_lookups['levels']
    if 'YoEd' in entry and pd.notna(entry['YoEd'])
}

df_students['ExpectedAge'] = df_students['Grade Level'].map(official_age_map)

# Step 3: Categorize AgeGroup based on reported Age vs. ExpectedAge
df_students['Age'] = pd.to_numeric(df_students['Age'], errors='coerce')

df_students['AgeGroup'] = df_students.apply(
    lambda row: (
        'Official Age' if pd.notna(row['Age']) and pd.notna(row['ExpectedAge']) and row['Age'] == row['ExpectedAge']
        else 'Under Age' if pd.notna(row['Age']) and pd.notna(row['ExpectedAge']) and row['Age'] < row['ExpectedAge']
        else 'Over Age' if pd.notna(row['Age']) and pd.notna(row['ExpectedAge']) and row['Age'] > row['ExpectedAge']
        else 'Unknown'
    ),
    axis=1
)

# Step 4: Group and calculate percentages
age_group_dist = df_students[
    df_students['Grade Level'].notna() & df_students['AgeGroup'].notna()
].groupby(['Grade Level', 'AgeGroup']).size().unstack(fill_value=0)

# Reorder grades if necessary
grade_order_clean = [entry['N'] for entry in core_lookups['levels'] if entry['N'] in age_group_dist.index]
age_group_dist = age_group_dist.reindex(grade_order_clean)

# Convert counts to row-wise percentages
age_group_pct = age_group_dist.div(age_group_dist.sum(axis=1), axis=0) * 100

# Plot 1: All age groups (stacked %)
age_group_pct.plot(
    kind='bar', stacked=True, figsize=(12, 6),
    title='(Official Cleaned Data) % Enrollment by Grade and Age Group',
    ylabel='Percentage of Students', xlabel='Grade Level',
    rot=45, grid=True, legend=True
)
plt.tight_layout()
plt.show()

# Plot 2: Only Under and Over Age (stacked %)
cols_to_plot = [col for col in age_group_pct.columns if col in ['Under Age', 'Over Age']]
age_group_pct[cols_to_plot].plot(
    kind='bar', stacked=True, figsize=(12, 6),
    title='(Official Cleaned Data) % Enrollment by Grade: Under & Over Age Only',
    ylabel='Percentage of Students', xlabel='Grade Level',
    rot=45, grid=True, legend=True
)
plt.tight_layout()
plt.show()


# %%
