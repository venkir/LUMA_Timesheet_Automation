
# Connect to Outlook using Microsoft Graph API and fetch meetings with "LUMA" in the title
# Authenticate with Microsoft Graph
# Filter meetings by title
# Extract date, duration, and agenda
# Write to CSV with static fields

import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
import win32com.client
from msal import ConfidentialClientApplication

# === Static Info ===
EMPLOYEE_NAME = "Venki Ramachandran"
IDENTIFIER = "CG11"
POSITION = "Industry Subject Matter Expert (SMS)"
PROJECT_ID = "14F000000000"
TASK_CODE = "A3403"
CLIENT_KEYWORD = "LUMA Timesheet Entry"

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)  # 9 = Calendar

# Define time range
now = datetime.now(timezone.utc)

start_str = "08/25/2025" # Change this for each run, typically a Monday
end_str = "08/30/2025"   # Change this for each run, typically a Saturday
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
            "Site": "OffIsland",
            "Hours": round(duration, 2),
            "PROJECT_ID": PROJECT_ID,
            "Task Code": TASK_CODE, 
            "Task Description": appt.Body.strip()
        })         
    except Exception as e:
        print(f"⚠️ Skipped an appointment due to error: {e}")


# Export to CSV
df = pd.DataFrame(timesheet)

if df.empty:
    err_msg = f"❌ Error: No meetings found with title '{CLIENT_KEYWORD}' from: {start_str} to {end_str}. Please check your Outlook calendar and try again."
    print(err_msg)
else:
    df.to_csv("timesheet_luma.csv", index=False)
    print("✅ Timesheet exported to timesheet_luma.csv")