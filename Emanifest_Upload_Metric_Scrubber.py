import pandas as pd
import os

# Emanifest Upload metric Data scrubbing
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

# Concatenate the dataframes
master_dv = pd.concat(dataframes, axis=0)
master_dv.columns = ["Accession Number", "Project ID", "User", "Completion Date"]

# Convert "Completion Date" to datetime format
# master_dv["Completion Date"] = pd.to_datetime(master_dv["Completion Date"]).dt.strftime('%m/%d/%Y')

# Identify rows with unique dates
unique_dates = master_dv.drop_duplicates(subset=['Project ID' , 'Completion Date'])

print(unique_dates)

# Save unique dates to Excel file
unique_dates.to_excel('/Users/gsikora/Desktop/GH/5-Data_Metrics/EMAN/2023/Master File/master_EMAN_metrics.xlsx', sheet_name='EMAN Metrics', index=False)
