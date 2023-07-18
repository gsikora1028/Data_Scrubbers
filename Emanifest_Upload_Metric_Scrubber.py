# ********************************************************************************************
# Program: E-Manifest Metric Scrubber
# Author: Gabe Sikora
# Date: 5/2/2023
# Function: Concatenate EMAN Adhoc Query Files into a single .xlsx file --> Pass to Tableau
# ********************************************************************************************

import pandas as pd
import os

print ('Accessing E-Manifest Adhoc Query Files...')

# List EMAN Adhoc Query filepaths
filepaths = [
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/1-Jan-EMAN.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/2-Feb-EMAN.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/3-Mar-EMAN.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/4-Apr1-EMAN.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/4-Apr2-EMAN.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/5-May-EMAN.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/6-Jun-EMAN.xlsx'
]

# Read all the Excel files into separate dataframes
dataframes = [pd.read_excel(filepath) for filepath in filepaths]

# Concatenate the dataframes into 1 master dataframe
master_dv = pd.concat(dataframes, axis=0)
master_dv.columns = ["Accession Number", "Project ID", "User", "Completion Date"]

# Identify rows with unique dates
unique_dates = master_dv.drop_duplicates(subset=['Project ID' , 'Completion Date'])
print(unique_dates)

# Save unique dates to Excel file (indicates the number of manifests uploaded)
unique_dates.to_excel('/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/Master File/master_EMAN_metrics.xlsx', sheet_name='EMAN Metrics', index=False)
