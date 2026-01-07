
# ---------------------------------------------------------------------------
# GetSharepointFilesInventory.py
# Created on: 2024-02-05
# Created by: Crystal Piroozy Posey
# Description: 
# An automated process to access documents stored in SharePoint and generate a comprehensive inventory of these files. 
# The desired output includes two formats: an Excel spreadsheet and a text file. 
# The goal is to efficiently catalog the documents within SharePoint, providing a detailed record that can be easily referenced and shared. 
# This process aims to streamline document management and enhance accessibility for stakeholders by consolidating file information into organized and easily readable formats.
# ---------------------------------------------------------------------------

#packages utilized in script
from shareplum import Site, Office365
from requests.auth import HTTPBasicAuth
import getpass
import xlwt
from datetime import datetime
from xlwt import Workbook
from requests_ntlm import HttpNtlmAuth
from shareplum.site import Version
from office365.sharepoint.client_context import ClientContext
import openpyxl.styles as styles
import os
import pandas as pd

print("Welcome to the SharePoint Document Inventory Script. This tool facilitates logging into a SharePoint site and retrieving an inventory of document files within a specified project site. The excel file will output where you specify...")
print("Note: If your organization utilizes Multi-Factor Authentication (MFA) or has a stringent security sign-in process, access may be denied. In such cases, please contact your IT department for guidance on registering applications in Azure AD.")

print("\nPlease enter the URL of the SharePoint site. For example, use a format similar to this (though not necessarily identical):")
print("https://organization.sharepoint.com/sites/nameofproject")

print("Site URL:")
site_url = input()

print("To proceed, kindly enter your username. Typically, this is the email associated with your SharePoint site:")
username = input()
print("Enter your password (the password is hidden, you will not see any characters typing)--")
password = str(getpass.getpass())

print("Please enter the full path to where you want the files saved (no quotations needed): ")
cwd = input()

print("SharePoint Inventory Script is starting...")
ctx = ClientContext(site_url).with_user_credentials(username, password)
lists = ctx.web.get_folder_by_server_relative_url("Shared Documents")
current_path = "Shared Documents"

wb = Workbook()
excel_file = wb.add_sheet('SharePoint File Directory')

#functions for execution start below--



def write_to_excel(row, col, value):
    # Directly write to the worksheet (excel_file is the worksheet)
    excel_file.write(row, col, value)

def process_files(folder, current_path, row, col):
    folder_path = current_path + "/{0}".format(folder)
    folder_lists = ctx.web.get_folder_by_server_relative_url(folder_path)
    file_names = folder_lists.files
    ctx.load(file_names)
    ctx.execute_query()

    print("#########################################################")
    print("Folder: " + folder)
    write_to_excel(row, col, folder)
    write_to_excel(row, col+1, "Folder")
    col=0
    row += 1
    col=0

    for item in file_names:
        print("File name: {0}".format(item.properties['Name']))
        write_to_excel(row, col, item.properties['Name'])
        write_to_excel(row, col + 1, "File")
        
        write_to_excel(row, col + 2, item.properties['LinkingUrl'])
        
        
        # Retrieve TimeCreated and TimeLastModified from subfile properties
        time_created = item.properties.get("TimeCreated", datetime.min)
        time_last_modified = item.properties.get("TimeLastModified", datetime.min)

        # Convert to MM/DD/YYYY format
        formatted_time_created = time_created.strftime("%m/%d/%Y")
        formatted_time_last_modified = time_last_modified.strftime("%m/%d/%Y")

        # Write to Excel
        write_to_excel(row, col + 3, formatted_time_created)
        write_to_excel(row, col + 4, formatted_time_last_modified)
        
        
        
        write_to_excel(row, col + 5, item.properties['Length'])
        col=0
        
        row += 1

        col=0
    row = process_subfolders(folder, folder_lists, folder_path, row, col)

    return row

def process_subfolders(subfolder, sublists, current_path, row, col):
    subfolder_names = sublists.folders
    ctx.load(subfolder_names)
    ctx.execute_query()

    for subfolder_item in subfolder_names:
        print("-------------------------------------------")
        print("\tSubFolder under {}".format(subfolder_item.properties["Name"]))
        col=0
        write_to_excel(row, col, subfolder_item.properties["Name"])
        write_to_excel(row, col+1, "Folder")
        row += 1

        sub_subfolder_path = current_path + "/{0}".format(subfolder_item)
        subfile_lists = ctx.web.get_folder_by_server_relative_url(sub_subfolder_path)
        subfiles_names = subfile_lists.files
        ctx.load(subfiles_names)
        ctx.execute_query()

        for subfile in subfiles_names:
            print("\t\tFile name: {0}".format(subfile.properties["Name"]))
            write_to_excel(row, col, subfile.properties["Name"])
            write_to_excel(row, col + 1, "File")
            write_to_excel(row, col + 2, subfile.properties['LinkingUrl'])
             # Retrieve TimeCreated and TimeLastModified from subfile properties
            time_created = subfile.properties.get("TimeCreated", datetime.min)
            time_last_modified = subfile.properties.get("TimeLastModified", datetime.min)

            # Convert to MM/DD/YYYY format
            formatted_time_created = time_created.strftime("%m/%d/%Y")
            formatted_time_last_modified = time_last_modified.strftime("%m/%d/%Y")

            # Write to Excel
            write_to_excel(row, col + 3, formatted_time_created)
            write_to_excel(row, col + 4, formatted_time_last_modified)
            
            write_to_excel(row, col + 5, subfile.properties['Length'])
            col=0
            row += 1
            col=0

        row = process_subfolders(subfolder_item, subfile_lists, sub_subfolder_path, row, col + 1)

    return row

def getDocumentStatus(sharepoint_file):
    excel_file = r'C:\Users\CPosey\Downloads\query (1).xlsx' 
    df_excel = pd.read_excel(excel_file)
    df_sharepoint = pd.read_excel(sharepoint_file)

    name = 'Name'
    doc_status = 'Doc Status'
    for index, value in df_excel[name].items():
        name_value = value
        doc_status_value = df_excel.at[index, doc_status]
        for index, value in df_sharepoint["Name"].items():
            if value == name_value:
                
                
            
        
        
      

def main():
    folder_names = lists.folders
    ctx.load(folder_names)
    ctx.execute_query()

    row = 0
    col = 0
    
    #headers
    excel_file.write(row, col, "Name")
    excel_file.write(row, col+1, "Type")
    excel_file.write(row, col+2, "Linking URL")
    excel_file.write(row, col+3, "Time Created")
    excel_file.write(row, col+4, "Time Modified")
    excel_file.write(row, col+5, "Length/Size (bytes)")
    excel_file.write(row, col+6, "Doc Status")
    row+=1
    

    for item in folder_names:
        row = process_files(item.properties["Name"], current_path, row, col)
        row += 1
        col=0


    
    try:
        
        # Save Excel file with raw string
        excel_filename = r'SharePointInventoryOutput.xls'
        cwd_excel = os.path.join(cwd, excel_filename)
        wb.save(cwd_excel)
        print(f"Excel file saved successfully: {cwd_excel}")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
    
#end of script