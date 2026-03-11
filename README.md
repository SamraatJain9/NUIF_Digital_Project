# NUIF Digital Projects

# HNM - Human Network Maintainer
This repository contains a Google Apps Script file to automate contact management and reminder emails in a Google Sheet: **hnm.js**.

---

## hnm.js

**Purpose:**  
This script powers a personal networking workflow in Google Sheets by:

- Creating and formatting a standard contact sheet layout.
- Tracking key reminder triggers (birthdays, anniversaries, and follow-up intervals).
- Sending a daily email digest of contacts that need attention.
- Providing menu actions to run reminders and manage automation.

**Columns Created (A -> Q):**

| Name | Email | Phone Number | LinkedIn | Company | Title | Industry | Country of Residence | Religion | Birthday | Holidays | Last Meeting | Contact Interval | Anniversary | (empty) | Recipient Email | Trigger hour (0-23) |
|------|-------|--------------|----------|---------|-------|----------|----------------------|----------|----------|----------|--------------|------------------|-------------|--------|-----------------|---------------------|

**Features:**

- `setupSheet()`:
  - Creates required headers and formatting (bold white text, blue background, central alignment).
  - Freezes row 1 and applies consistent column widths.
  - Applies date formatting to Birthday, Last Meeting, and Anniversary columns.
  - Adds dropdown validation for Contact Interval.
  - Sets defaults for Recipient Email (P2) and Trigger Hour (Q2).
- `sendReminders()`:
  - Scans all contact rows.
  - Detects today's Birthday and Anniversary matches.
  - Detects Follow up reminders from Last Meeting + Contact Interval.
  - Sends one HTML table email digest via `MailApp.sendEmail`.
- `setupDailyTrigger()`:
  - Reads trigger hour from Q2.
  - Replaces old `sendReminders` triggers.
  - Creates a daily time-based trigger.
- `removeAllTriggers()`:
  - Deletes all project triggers.
- `onOpen()`:
  - Adds **Setup** and **Reminders** menus to the Google Sheet UI.

**Usage:**

1. Open your Google Sheet and go to **Extensions -> Apps Script**.
2. Paste `personal_networking.js` into the editor and save.
3. Reload the spreadsheet.
4. From **Setup -> Setup sheet**, initialize the sheet structure.
5. Optionally run **Reminders -> Set daily reminder** to automate notifications.
6. Use **Reminders -> Run reminders now** to test immediately.

**Notes**

- Reminder matching is based on current day/month checks for date triggers.
- Follow-up reminders depend on valid values in **Last Meeting** and **Contact Interval**.
- If Q2 is empty or invalid, daily reminders default to hour `9` (09:00).
- Running setup multiple times rewrites headers/formatting but preserves data rows below row 1.

### Authors
[Samraat Jain](https://github.com/SamraatJain9), [James Delin](https://github.com/jd-0001), [Sarah Rafiepour](https://github.com/sarahr15), [Ryan Duong](https://github.com/RyanDuong0), [Shalom Ademuwagun](https://github.com/ChachyDev)
