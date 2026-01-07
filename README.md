# recursive-structure-mapper
This Python tool recursively traverses SharePoint file structures to map and inventory all files and folders. It uses recursive functions to efficiently navigate nested directories, collect metadata, and produce a comprehensive inventory report for auditing or analysis.
Description:
An automated process to access documents stored in SharePoint and generate a comprehensive inventory of these files.
The desired output includes two formats: an Excel spreadsheet and a text file.
The goal is to efficiently catalog the documents within SharePoint, providing a detailed record that can be easily referenced and shared.
This process aims to streamline document management and enhance accessibility for stakeholders by consolidating file information into organized and easily readable formats.

Usage:
1. Run the script and provide SharePoint site URL, username, password, and the desired output path.
2. The script will generate an Excel spreadsheet containing a detailed inventory of documents.

Dependencies:
- shareplum
- requests_ntlm
- office365

Note:
If your organization utilizes Multi-Factor Authentication (MFA) or has a stringent security sign-in process, access may be denied.
Contact your IT department for guidance on registering applications in Azure AD.
