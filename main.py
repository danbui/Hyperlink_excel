import openpyxl

# Define the path to your Excel file
excel_file_path = r'C:\Users\Administrator\Desktop\Mother_SKU.xlsx'
# Define the column that contains the folder names (e.g., 'A' for column A)
folder_name_column = 'A'
# Define the base URL of your Google Drive folder
google_drive_base_url = 'https://onedrive.live.com/?id=root&cid=4D867D3DAA155D5E'

# Load the workbook and select the active worksheet
workbook = openpyxl.load_workbook(excel_file_path)
worksheet = workbook.active

# Iterate over the cells in the specified column
for row in range(2, worksheet.max_row + 1):  # Assuming the first row is a header
    cell_value = worksheet[f'{folder_name_column}{row}'].value
    if cell_value:
        # Create the hyperlink
        folder_url = google_drive_base_url + cell_value
        worksheet[f'{folder_name_column}{row}'].hyperlink = folder_url
        worksheet[f'{folder_name_column}{row}'].style = 'Hyperlink'

# Save the workbook
workbook.save(excel_file_path)

print("Hyperlinks added successfully.")
