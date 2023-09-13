import pyodbc
import pandas as pd
import datetime
import os

# Connect to SQL Server
server = 'uat-synapse-workspace.sql.azuresynapse.net'
database = 'UAT_abs_ent_edw'
authentication = 'ActiveDirectoryIntegrated'

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';Authentication=' + authentication)

# Define your queries
current_quarter = (datetime.datetime.now().month - 1) // 3 + 1
timecard_query = f'''
SELECT per.employee_number as 'Employee Number', per.last_name as 'Last Name', CONVERT(Date, expenditure_item_date) as 'Trans Date', p.project_status_code as 'Status',
p.proj_number as 'Project Number', p.proj_name as 'Work Order\Proj', p.project_description as 'Description/Vessel',
t.expenditure_type as 'Expenditure Type', t.bureau_regular_hours as 'Hours',
t.bureau_utilization_group as 'Hour Type'--, task_number
FROM gold.dim_project p JOIN gold.fact_timecard t on p.pk_project = t.fk_project
JOIN gold.dim_personnel per on t.incurred_by_person_id = per.person_id
WHERE per.last_name IN ('Fukuda', 'Hirai', 'Okamoto', 'Yoshida', 'Naito', 'Ookawa', 'Morioka', 'Megahed')
AND CONVERT(Date, expenditure_item_date) > '2023-01-01'
;'''

# Read data into a data frame
timecard_df = pd.read_sql(timecard_query, conn)

# New data frames for excel work sheets
# Port DFs - Filters 
port_df = timecard_df[['Employee Number', 'Last Name', 'Trans Date','Status','Project Number']]
port_df_pivot = timecard_df[['Employee Number', 'Last Name', 'Trans Date']].drop_duplicates(subset='Trans Date')[['Last Name', 'Employee Number', 'Trans Date']]

# NC Days DFs
nc_days_df = port_df_pivot[['Employee Number', 'Last Name', 'Trans Date']].groupby(['Employee Number', 'Last Name']).count().reset_index()[['Last Name', 'Employee Number', 'Trans Date']]
nc_days_df = nc_days_df.rename(columns={'Trans Date': 'Count of Trans Date'})

# Display Data
print(nc_days_df)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(r'C:\Users\ZDenton\Documents\EBI - 240 (python)\output.xlsx', engine='xlsxwriter')

# Write each dataframe to worksheet.
# Timecard Detail
timecard_df.to_excel(writer, sheet_name='Timecard Detail', index=False)
# Port
port_df.to_excel(writer, sheet_name='Port', index=False)
port_df_pivot.to_excel(writer, sheet_name='Port Pivot', index=False)

# NC Days
port_df_pivot.to_excel(writer, sheet_name='NC Days', index=False)
nc_days_df.to_excel(writer, sheet_name='NC Days Pivot', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()




    


      
