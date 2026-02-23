# NUIF Digital Projects

# Rolodex Reminders - Google Sheets Automation

This repository contains two Google Apps Script files to automate reminders for contacts in a Google Sheet:

1. **rolodex.js** – Handles sending automated reminders for birthdays, anniversaries, and touch intervals.
2. **sheetSetup.js** – Provides a one-click setup for the Google Sheet, creating and formatting all required columns.

---

## 1. rolodex.js

**Purpose:**  
This script scans your contact sheet and sends reminder emails based on the following triggers:

- **Birthday** – Matches contacts whose birthday falls on the current day.
- **Anniversary** – Matches contacts whose anniversary falls on the current day.
- **Touch Interval** – Matches contacts whose last interaction plus the touch interval (in quarters) falls on the current day.

**Features:**

- Supports batching for large sheets (up to 900 rows per batch).
- Handles timezone correctly using the spreadsheet’s timezone.
- Sends a single email with all reminders in a formatted table.
- Daily automated trigger can be configured via `setupDailyTrigger()`.
- Includes a menu in Google Sheets for manual execution (`Run reminders now`) or trigger setup.

**Usage:**

1. Paste `rolodex.js` into the Apps Script editor attached to your Google Sheet.
2. Reload the sheet. You will see a **🔔 Reminders** menu.
3. Optionally, run **Set up daily trigger** to send reminders automatically every day.
4. Run **Run reminders now** for immediate testing.

---

## 2. sheetSetup.js

**Purpose:**  
Sets up a new Google Sheet with all required columns in the correct order, pre-filled defaults, and formatted headers. This allows users to start using `rolodex.js` without manually configuring the sheet.

**Columns Created (A → U):**

| Name | Email | Phone Number | LinkedIn | Company | Title | Industry | Country of Residence | City | Timezone | Religion | Birthday | Holidays | Last Interaction | Last Meeting | Touch Interval (Quater) | Last Conversation Notes | Anniversary | (empty) | Recipient Email | Trigger hour (0–23) |
|------|-------|--------------|----------|---------|-------|----------|--------------------|------|----------|---------|---------|---------|-----------------|--------------|------------------------|------------------------|------------|--------|----------------|--------------------|

**Features:**

- Creates and formats the header row:
  - Navy blue background
  - White bold text
  - Centered alignment
- Freezes the header row
- Sets default **Recipient Email** (T2) and **Trigger Hour** (U2)
- Resizes columns for better readability

**Usage:**

1. Paste `sheetSetup.js` into the Apps Script editor attached to your Google Sheet.
2. Add the following `onOpen()` function to show a menu button:

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔔 Reminders")
    .addItem("Setup sheet", "setupSheet")
    .addToUi();
}
```
3. Reload the Google Sheet.
4. Click 🔔 Reminders → Setup sheet to create and format all columns.
5. The sheet is now ready for use with rolodex.js.

**Notes**

- The scripts rely on the spreadsheet timezone to determine “today,” so all triggers and date calculations are consistent regardless of the user’s local timezone.
- rolodex.js is safe for large datasets due to batching and caching.
- sheetSetup.js is idempotent: running it multiple times will overwrite the header row without affecting existing data below row 1.

### Authors
[Samraat Jain](https://github.com/SamraatJain9), [James Delin](https://github.com/jd-0001), [Sarah Rafiepour](https://github.com/sarahr15), [Ryan Duong](https://github.com/RyanDuong0), [Shalom Ademuwagun](https://github.com/ChachyDev)