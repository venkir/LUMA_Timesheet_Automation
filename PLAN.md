# Project: Outlook Meeting Exporter for Timesheets

## Objective
Automatically extract Outlook calendar meetings related to client "LUMA" and generate timesheet entries with static and dynamic fields.

## Must-Haves
- Connect to Outlook calendar (Microsoft Graph API)
- Filter meetings with "LUMA" in the title
- Extract:
  - Date
  - Duration
  - Agenda/Description
- Generate timesheet entry with:
  - Name: Venki Ramachandran
  - Employee ID: [Your ID]
  - Date & Duration from meeting
  - Task Description from agenda

## Nice-to-Haves
- Export to CSV or Excel
- UI to trigger export manually
- Scheduled daily/weekly export

## Out of Scope
- Editing calendar events
- Integration with external timesheet systems (for now)