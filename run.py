import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import sys

# Load the Excel file
file_path = sys.argv[1]
workbook = load_workbook(filename=file_path)
sheet = workbook.active

# Convert the active sheet to a pandas DataFrame
data = sheet.values
columns = next(data)
df = pd.DataFrame(data, columns=columns)

# Filter the dataframe based on the given criteria
filtered_df = df[df['VIN-Ready for pick up'].str.lower() == 'yes']
filtered_df = filtered_df[filtered_df['Term File Submitted'].isna()]

# Calculate the days between 'Date Created' and today
filtered_df['Days Between'] = (datetime.now() - pd.to_datetime(filtered_df['Date Created'])).dt.days

# Select the desired columns
final_report = filtered_df[['VIN', 'Vehicle Type', 'Station', 'Address', 'City', 'State', 'ZIP', 'Provider', 'Days Between']]
final_report.rename(columns={'ZIP': 'Zip'}, inplace=True)

# Save the final report to a new Excel file
output_path = 'Final_Report.xlsx'
final_report.to_excel(output_path, index=False)

print("Final report has been generated and saved to", output_path)