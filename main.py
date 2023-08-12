# Import necessary libraries
import pandas as pd
import json
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# Open and read the content of the Installments.json file
with open('Installments.json', 'r', encoding='utf-8') as file:
    data = json.load(file)

# Convert the JSON data to a pandas DataFrame for easier processing
df_main = pd.json_normalize(data)

# Dictionary for renaming specific columns to more human-readable names
column_renaming = {...}  # [truncated for brevity]

# Rename the columns in the DataFrame based on the dictionary
df_main.rename(columns=column_renaming, inplace=True)

# List of columns to be removed from the DataFrame
columns_to_drop = [...]  # [truncated for brevity]

# Drop the specified columns if they exist in the DataFrame
df_main = df_main.drop(columns=[col for col in columns_to_drop if col in df_main.columns])

# Dictionary to map messenger types to their proper names
messenger_mapping = {...}  # [truncated for brevity]

# Replace values in the 'Internal Messenger Type' column based on the mapping dictionary
df_main['Internal Messenger Type'] = df_main['Internal Messenger Type'].replace(messenger_mapping)

# Capitalize the values in the "Status" column for consistency
df_main['Status'] = df_main['Status'].str.capitalize()

# Capitalize the values in the "Gender" column for consistency
df_main['Gender'] = df_main['Gender'].str.capitalize()

# Save the processed DataFrame to an Excel file
output_file = 'output.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_main.to_excel(writer, sheet_name='Main', index=False)

# Load the saved Excel file back into memory for formatting
wb = openpyxl.load_workbook(output_file)

# Define styles for formatting cells in the Excel sheet
font = Font(name='Vazirmatn')
alignment = Alignment(horizontal='center', vertical='center')
light_yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
light_green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

# Apply the defined styles to cells in the Excel sheet
sheet = wb['Main']
for row in sheet:
    for cell in row:
        cell.font = font
        cell.alignment = alignment
        cell.number_format = '@'  # Set number format to text
        # Differentiate the header row with a unique background color
        if cell.row == 1:
            cell.fill = light_yellow_fill
        else:
            cell.fill = light_green_fill

    # Adjust the width of each column based on the max length of its content
    for column_cells in sheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length

# Save the formatted Excel file
wb.save(output_file)
