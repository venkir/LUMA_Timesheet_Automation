# Connect to Outlook using Microsoft Graph API and fetch meetings with "LUMA" in the title
# Authenticate with Microsoft Graph
# Filter meetings by title
# Extract date, duration, and agenda
# Write to CSV with static fields
# Venki: 09/23/25: SharePoint Integration done. Go to 'Timesheet ..' folder in SharePoint folder, click on 3 dots, click on Sync
#                : Added the very long common meeting xlsx file with spaces and double back slashes in the config.ini file
#                : Remove rows where the concatenated field 'Task Description' from Common Meeting List is empty after appending
#                : Input Dates are 4 months apart, so fetch all meetings in that range and changed filter to provide top 1000 meetings
#                : Added another parameter to the config file to create the final timesheet csv file in the same folder as the common meeting list file in SharePoint
import configparser
import requests
import msal
import pandas as pd
from datetime import datetime, timedelta, timezone
import json
import os
import sys

from bs4 import BeautifulSoup, Comment


def check_if_file_open(file_path):
    try:
        # Try opening the file in append mode
        with open(file_path, 'a'):
            pass
        print("File Creation is possible, Writing...")
    except PermissionError:
        print(f"File '{file_path}' appears to be open or locked by another process. Exiting")
        sys.exit(1)  # Exit gracefully
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

# Load the configuration file
config = configparser.ConfigParser()
config.read('config.ini')

# Outlook Access values
DISPLAY_NAME = config['Outlook']['DISPLAY_NAME']
CLIENT_ID = config['Outlook']['CLIENT_ID']
OBJECT_ID = config['Outlook']['OBJECT_ID']
TENANT_ID = config['Outlook']['TENANT_ID']
CLIENT_SECRET = config['Outlook']['SECRET_KEY']
SECRET_ID = config['Outlook']['SECRET_ID']
EXPIRY_DATE = config['Outlook']['EXPIRY_DATE']
meetings_file_name = config['Files']['meetings_file_name'].strip('"')

# Personal Info values
EMPLOYEE_NAME = config['Personal']['EMPLOYEE_NAME']
IDENTIFIER = config['Personal']['IDENTIFIER']   
POSITION = config['Personal']['POSITION']
PROJECT_ID = config['Personal']['PROJECT_ID']
TASK_CODE = config['Personal']['TASK_CODE']
CLIENT_KEYWORD = config['Personal']['CLIENT_KEYWORD']
SITE = config['Personal']['SITE']
user_id = config['Personal']['EMAIL_ID']
# === Static Info ===
STATIC_TEXT = "This is  in support to the Project implementation for the Advanced Metering Infrastructure (AMI) - Technical Integration Services for Outage Management System (OMS), \
    Geographic Information System (GIS), and Customer Emergency Management System (CEMS) as per contract 2025-L00157 related to Request for Proposal 183353"

#print(f"Client id : {CLIENT_ID}")

from azure.identity import DeviceCodeCredential
from msgraph import GraphServiceClient


# Microsoft Graph endpoints
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPE = ['https://graph.microsoft.com/.default']
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# Create a confidential client
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

# Acquire token
token_response = app.acquire_token_for_client(scopes=SCOPE)
access_token = token_response.get('access_token')

# Define headers
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# Define time range
now = datetime.now(timezone.utc)

start_str = "06/01/2025" # Change this for each run, typically a Monday
end_str = "09/30/2025"   # Change this for each run, typically a Saturday
# Example: For week of Aug 26 to Aug 31, 2024   
print(f"Fetching calendar events from {start_str} to {end_str}")

# Convert to ISO 8601 format with UTC timezone
start_date = datetime.strptime(start_str, "%m/%d/%Y").replace(tzinfo=timezone.utc).isoformat().replace("+00:00", "Z")
end_date = datetime.strptime(end_str, "%m/%d/%Y").replace(tzinfo=timezone.utc).isoformat().replace("+00:00", "Z")

# Build the URL correctly
url = f"{GRAPH_ENDPOINT}/users/{user_id}/calendarView?startDateTime={start_date}&endDateTime={end_date}"
params = {
    '$orderby': 'start/dateTime',
    '$top': 1000
}
response = requests.get(url, headers=headers, params=params)

# Process filtered appointments
timesheet = []
max_items = 100  # Safety limit
count = 0

if response.status_code == 200:
    events = response.json().get('value', [])
    #filtered_events = [event for event in events if event['subject'].strip().lower() == CLIENT_KEYWORD.lower()]
    filtered_events = [event for event in events if event['subject'].strip() == 'LUMA Timesheet Entry']
    for event in filtered_events:
        start_time_str = event['start']['dateTime']
        end_time_str = event['end']['dateTime']
        # Convert to datetime objects
        start_time = datetime.fromisoformat(start_time_str)
        end_time = datetime.fromisoformat(end_time_str)

        duration = (end_time - start_time).total_seconds() / 3600
        # Extract agenda from event body
        html_content = event.get('body', {}).get('content', '')

        # Parse and clean HTML
        soup = BeautifulSoup(html_content, 'html.parser')

        # Remove style and script tags
        for tag in soup(['style', 'script']):
            tag.decompose()

        # Remove HTML comments
        for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
            comment.extract()

        # Get clean text
        agenda = soup.get_text(separator=' ', strip=True)

        timesheet.append({
            "Name": EMPLOYEE_NAME,
            "Identifier": IDENTIFIER,
            "Position": POSITION,
            "Date": start_time.date().isoformat(),
            "Site": SITE,
            "Hours": round(duration, 2),
            "PROJECT_ID": PROJECT_ID,
            "Task Code": TASK_CODE, 
            "Task Description": agenda + " " + STATIC_TEXT
        })
else:
    print(f"Error: {response.status_code}")
    print(response.text)

# Export to CSV
timesheet_df = pd.DataFrame(timesheet)
# Check if we got any appointments from Outlook
if timesheet_df.empty:
    err_msg = f"❌ Error: No Events found in your Calendar with title '{CLIENT_KEYWORD}' from: {start_str} to {end_str}. Please check."
    print(err_msg)
    sys.exit()
else:
    print(f"Total appointments found in date range: {len(timesheet_df)}")


# Read the file and print locally
meetings_df = pd.read_excel(meetings_file_name, usecols=["Date", "Hours", "Concate of all Required fields"], sheet_name="Meeting List", engine="openpyxl")
#print("Printing the top few lines from the meeting list file:")
if meetings_df.empty:
    err_msg = f"❌ Error: No Common meeting List File: {meetings_file_name} found. Please check its location and name."
    print(err_msg)
    sys.exit()

# Since the neetings list has onlt three columsn we are inbterested in and the rest are our personal details
# Rename the last column to Task Description
meetings_df.rename(columns={"Concate of all Required fields": "Task Description"}, inplace=True)
#print(meetings_df.head())
# 
# Ensure df2 has all columns of df1
meetings_df_aligned = meetings_df.reindex(columns=timesheet_df.columns)
# Strip time component from Date column
meetings_df_aligned['Date'] = meetings_df_aligned['Date'].dt.date
    
# Append the two DataFrames
result = pd.concat([timesheet_df, meetings_df_aligned], ignore_index=True)
# Fill the NaN values with the default values
result['Name'] = result['Name'].fillna(EMPLOYEE_NAME)
result['Identifier'] = result['Identifier'].fillna(IDENTIFIER)
result['Position'] = result['Position'].fillna(POSITION)
result['Site'] = result['Site'].fillna(SITE)
result['PROJECT_ID'] = result['PROJECT_ID'].fillna(PROJECT_ID)
result['Task Code'] = result['Task Code'].fillna(TASK_CODE)
# Drop all rows where Date or Task Description is NaN
result = result.dropna(subset=['Task Description'])
#print(result.head(10))

# Create the output CSV file path
output_csv_dir = config['Files']['output_csv_dir'].strip('"')
output_csv_file_name = os.path.join(output_csv_dir, EMPLOYEE_NAME.replace(" ", "_") + "_Timesheet_3rd_Invoice.csv")

# Check if the file is open before writing or exit as this point
check_if_file_open(output_csv_file_name)

# try to write the file
try:
    result.to_csv(output_csv_file_name, index=False)
    if os.path.exists(output_csv_file_name):
        print(f"✅ File '{output_csv_file_name}' created successfully.")
except Exception as e:
    print(f"❌ Failed to export timesheet: {e}")
