import os
import openpyxl

# Prompt the user to enter the file location of the Excel file
excel_file_location = input("Enter the file location of the Excel file: ")

# Prompt the user to enter the root folder
root_folder = input("Enter the root folder: ")

workbook = openpyxl.load_workbook(excel_file_location)
worksheet = workbook.active

# Create Projects folder if it doesn't exist
os.makedirs(root_folder, exist_ok=True)

# Iterate through rows and create folders
for i, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=1):
    project_name = row[0]
    project_folder_name = f"{i}. {project_name}"
    project_folder = os.path.join(root_folder, project_folder_name)
    os.makedirs(project_folder, exist_ok=True)
    os.makedirs(os.path.join(project_folder, "Bills for MyLeo"), exist_ok=True)
    os.makedirs(os.path.join(project_folder, "Bills Document"), exist_ok=True)
