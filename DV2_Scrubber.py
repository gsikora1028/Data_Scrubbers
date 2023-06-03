import pandas as pd
import os

#DV2 Data scrubbing
filepaths = [
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/DV/2023/1-Jan-DV2.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/DV/2023/2-Feb-DV2.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/DV/2023/3-Mar-DV2.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/DV/2023/4-Apr-DV2.xlsx',
    '/Users/gsikora/Desktop/GH/5-Data_Metrics/DV/2023/5-May-DV2.xlsx'
]

# Read all the Excel files into separate dataframes
dataframes = [pd.read_excel(filepath) for filepath in filepaths]

master_dv = pd.concat(dataframes, axis=0)
master_dv.columns = ["Accession Number", "Project ID", "Test", "Type", "Completion Date", "User"]

# Convert "Completion Date" to datetime format
master_dv["Completion Date"] = pd.to_datetime(master_dv["Completion Date"]).dt.strftime('%m/%d/%Y')

master_dv.to_excel('/Users/gsikora/Desktop/GH/5-Data_Metrics/DV/2023/Master File/master_dv2_metrics.xlsx', sheet_name='DV2 Metrics', index=False)

print(master_dv)