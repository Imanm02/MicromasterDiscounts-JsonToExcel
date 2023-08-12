import pandas as pd
import json
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# Read the content of the Installments.json file
with open('Installments.json', 'r', encoding='utf-8') as file:
    data = json.load(file)

# Convert the main data to a DataFrame
df_main = pd.json_normalize(data)

# Rename the specified columns
column_renaming = {
    "created_at": "Created At",
    "updated_at": "Updated At",
    "status": "Status",
    "amount": "Amount",
    "discount_coefficient": "Discount Coefficient",
    "user.gender": "Gender",
    "user.name": "Name",
    "user.surname": "Surname",
    "user.email": "Email",
    "user.mobile_number": "Mobile Number",
    "user.national_code": "National Code",
    "user.phone_number": "Phone Number",
    "user.extra.telegram_number": "Telegram Number",
    "user.extra.whatsapp_number": "Whatsapp Number",
    "user.extra.internal_messenger_type": "Internal Messenger Type",
    "user.extra.internal_messenger_number": "Internal Messenger Number"
}

df_main.rename(columns=column_renaming, inplace=True)

# Drop the specified columns from df_main if they exist
columns_to_drop = [
    'user.extra.in_person_classes',
    'user.id',
    'user.photo',
    'user.bio',
    'user.created_at',
    'user.updated_at',
    'user.extra',
    'sources',  # Deleting the whole "sources" column
    'id',
    'user_id',
    'sharif_order_id',
    'reference_id'
]

df_main = df_main.drop(columns=[col for col in columns_to_drop if col in df_main.columns])

# Replace values in the 'Internal Messenger Type' column
messenger_mapping = {
    'bale': 'Bale',
    'eitaa': 'Eitaa',
    'rubika': 'Rubika',
    'gap': 'Gap',
    'soroush': 'Soroush',
    'igap': 'IGap'
}

df_main['Internal Messenger Type'] = df_main['Internal Messenger Type'].replace(messenger_mapping)

# Capitalize the "Status" column
df_main['Status'] = df_main['Status'].str.capitalize()

# Capitalize the "Status" column
df_main['Gender'] = df_main['Gender'].str.capitalize()

# Save to Excel
output_file = 'output.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_main.to_excel(writer, sheet_name='Main', index=False)

# Load the Excel file back into memory
wb = openpyxl.load_workbook(output_file)

# Define the styles to apply to the cells
font = Font(name='Vazirmatn')
alignment = Alignment(horizontal='center', vertical='center')
light_yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
light_green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

# Apply the styles to the cells
sheet = wb['Main']
for row in sheet:
    for cell in row:
        cell.font = font
        cell.alignment = alignment
        cell.number_format = '@'  # Set number format to text
        # Apply different background colors to the first row and the rest
        if cell.row == 1:
            cell.fill = light_yellow_fill
        else:
            cell.fill = light_green_fill

    # Adjust the width of the columns
    for column_cells in sheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length

# Save the changes made to the Excel file
wb.save(output_file)