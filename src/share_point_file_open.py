import os
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import openpyxl  # For opening Excel files
import ExcelInterface as EI

def open_file(file_path):
    # Check if file exists locally
    if os.path.exists(file_path):
        print(f"File found locally at: {file_path}")
        try:
            # Open the file (example for Excel file)
            wb = EI.openExcel(file_path)
            print("File opened successfully!")
            return wb
        except Exception as e:
            print(f"Failed to open the file: {e}")
    else:
        print(f"File not found locally. Attempting to fetch from SharePoint: {file_path}")
        # File is not local, try accessing it via SharePoint
        download_from_sharepoint(file_path)

def download_from_sharepoint(file_url):
    site_url = "https://your-sharepoint-site-url"  # Base SharePoint site URL
    client_id = "your-client-id"  # From Azure AD App Registration
    client_secret = "your-client-secret"  # From Azure AD App Registration

    # Authenticate with Client ID and Secret
    credentials = ClientCredential(client_id, client_secret)
    ctx = ClientContext(site_url).with_credentials(credentials)

    try:
        # Fetch the file from SharePoint
        response = File.open_binary(ctx, file_url)
        local_file_path = os.path.basename(file_url)  # Save with the same name
        with open(local_file_path, "wb") as local_file:
            local_file.write(response.content)
        print(f"File downloaded successfully from SharePoint and saved as {local_file_path}")
        # Attempt to open the file
        wb = openpyxl.load_workbook(local_file_path)
        print("File opened successfully after downloading!")
        return wb
    except Exception as e:
        print(f"An error occurred while fetching the file from SharePoint: {e}")

# Example Usage
file_path = input("Enter the file path: ")  # Dynamically provide path
open_file(file_path)
