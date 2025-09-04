# Instructions for AI Assistants

## Goal
Build a Python app that connects to Outlook via Microsoft Graph API, filters meetings with "LUMA" in the title, and generates timesheet entries.

## Key Tasks
1. Authenticate with Microsoft Graph API
2. Fetch calendar events
3. Filter by title containing "LUMA"
4. Extract:
   - Start time
   - End time
   - Agenda/Description
5. Format output:
   - Name: Venki Ramachandran
   - Employee ID: [Your ID]
   - Date
   - Duration
   - Task Description
## Output Format
CSV or Excel file with columns:
- Name
- Employee ID
- Date
- Duration
- Task Description
