
# Connect to Outlook using win32com client and fetch meetings with a known and fixed title
# Filter meetings by title
# Extract date, duration, and agenda
# Write to CSV with static fields
# 9/2/2025: V1.0 | Initial Draft
# 9/3/2025: V1.1 | Added error handling for missing meetings in the date range
# 9/4/2025: V1.2 | Added check for no meetings found, adding static text to all entries, checking if csv file is open and exiting gracefully, plus append the list of common meetings

import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
import win32com.client
from msal import ConfidentialClientApplication
import os
import sys

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

# === Static Info ===
EMPLOYEE_NAME = "Venki Ramachandran"
IDENTIFIER = "CG11"
POSITION = "Industry Subject Matter Expert (SMS)"
PROJECT_ID = "14F000000000"
TASK_CODE = "A3403"
CLIENT_KEYWORD = "LUMA Timesheet Entry"
SITE = "OffIsland"
STATIC_TEXT = "This is  in support to the Project implementation for the Advanced Metering Infrastructure (AMI) - Technical Integration Services for Outage Management System (OMS), \
    Geographic Information System (GIS), and Customer Emergency Management System (CEMS) as per contract 2025-L00157 related to Request for Proposal 183353"
MEETING_LIST_FILE_NAME = "https://capgemininar.sharepoint.com/:x:/r/sites/CGinternalLUMASmartMeterProject/_layouts/15/Doc.aspx?sourcedoc=%7B764391CB-55AC-4420-A938-A381937BD97C%7D&file=Venki%20Ramachandran%20Task%20Excel%20for%20Milestone%20Invoicing-3rd%20Invoice.xlsx&action=default&mobileredirect=true"

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)  # 9 = Calendar

# Define time range
now = datetime.now(timezone.utc)

start_str = "08/25/2025" # Change this for each run, typically a Monday
end_str = "08/29/2025"   # Change this for each run, typically a Saturday
# Example: For week of Aug 26 to Aug 31, 2024   

start = tart = datetime.strptime(start_str, "%m/%d/%Y").replace(tzinfo=timezone.utc)    # Change to desired start date
end = tart = datetime.strptime(end_str, "%m/%d/%Y").replace(tzinfo=timezone.utc)     # Change to desired end date


# Sort and restrict items
appointments = calendar.Items
appointments.Sort("[Start]")
appointments.IncludeRecurrences = True

# Restrict to time range
restriction = f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{end.strftime('%m/%d/%Y %H:%M %p')}'"
restricted_items = appointments.Restrict(restriction)

# Process filtered appointments
timesheet = []
max_items = 100  # Safety limit
count = 0

for appt in restricted_items:
    if count >= max_items:
        break
    count += 1

    try:
        if CLIENT_KEYWORD.lower() in appt.Subject.lower():
            start_time = appt.Start
            end_time = appt.End
            duration = (end_time - start_time).total_seconds() / 3600
            agenda = appt.Body.strip() if appt.Body else "No agenda provided"

            timesheet.append({
            "Name": EMPLOYEE_NAME,
            "Identifier": IDENTIFIER,
            "Position": POSITION,
            "Date": appt.Start.date().isoformat(),
            "Site": SITE,
            "Hours": round(duration, 2),
            "PROJECT_ID": PROJECT_ID,
            "Task Code": TASK_CODE, 
            "Task Description": appt.Body.strip() + " " + STATIC_TEXT
        })         
    except Exception as e:
        print(f"⚠️ Skipped an appointment due to error: {e}")

# Export to CSV
timesheet_df = pd.DataFrame(timesheet)

# Download from SharePoint and save it locally as Common_Meeting_List.xlsx
meetings_file_name = r"C:\Users\venramac\Downloads\Common_Meeting_List.xlsx"
# Read the file and print locally
meetings_df = pd.read_excel(meetings_file_name, usecols=["Date", "Hours", "Concate of all Required fields"], sheet_name="Meeting List", engine="openpyxl")
#print("Printing the top few lines from the meeting list file:")

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
print(result.head(10))

# Check if the file is open before writing
check_if_file_open("timesheet_luma.csv")

if timesheet_df.empty:
    err_msg = f"❌ Error: No meetings found with title '{CLIENT_KEYWORD}' from: {start_str} to {end_str}. Please check your Outlook calendar and try again."
    print(err_msg)
else:
    result.to_csv("timesheet_luma.csv", index=False)
    print("✅ Timesheet exported to timesheet_luma.csv")