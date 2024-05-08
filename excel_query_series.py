"""
Script Name: excel_query_series.py
Author: Natasha Seguin
Date: May 8, 2024
Description: This script calculates the sum of a field based on a series of queries listed in an excel spreadsheet.

Usage: Inputs must be excel files.

Dependencies: Python extension. Other extension(s) may be required if code is amended to use GeoPandas.

License: N/A

Version: 1.0

Contact:
    Email: Natasha.Seguin@gov.bc.ca
    GitHub: 

"""

# Import modules
import pandas as pd
import os

workspace_path = r"INSERT PATH"

# Read the Excel spreadsheet into a pandas DataFrame
excel_data = pd.read_excel(os.path.join(workspace_path, "INSERT INPUT EXCEL FILE NAME"))

# Read the Excel spreadsheet for exported attribute table into a DataFrame - 
attribute_data = pd.read_excel(os.path.join(workspace_path, "INSERT INPUT EXCEL FILE NAME"))

# Iterate through the rows of the first Excel DataFrame and perform queries on attribute dataset for each row
for index, row in excel_data.iterrows():
    # Extract values from the Excel row
    field1_value = row['INSERT COLUMN HEADER NAME']
    field2_value = row['INSERT COLUMN HEADER NAME'] 
    field3_value = row['INSERT COLUMN HEADER NAME']
    
    # Perform attribute query
    query_result = attribute_data[(attribute_data['INSERT FIELD NAME'] == field1_value) &
                                (attribute_data['INSERT FIELD NAME'] == field2_value) &
                                (attribute_data['INSERT FIELD NAME'] == field3_value)]
    
    # Perform calculations based on the query result and update the Excel DataFrame
    excel_data.at[index, 'INSERT NEW COLUMN HEADER NAME'] = query_result['INSERT FIELD NAME'].sum()  # Calculating the sum of area
    
# Write the updated DataFrame to a new Excel file
excel_data.to_excel(os.path.join(workspace_path, "INSERT OUTPUT EXCEL FILE NAME"), index=False)

print("Task complete")