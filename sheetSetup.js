/*
  Author(s): Samraat Jain, 
  Version: // A Rolodex Sheet Setup Script v1.0 — 2026-02-10
  This script:
  - Automatically sets up the sheet with the necessary headers and formatting for A Rolodex Script.
*/

function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var headers = [
    "Name",
    "Email",
    "Phone Number",
    "LinkedIn",
    "Company",
    "Title",
    "Industry",
    "Country of Residence",
    "City",
    "Timezone",
    "Religion",
    "Birthday",
    "Holidays",
    "Last Interaction",
    "Last Meeting",
    "Touch Interval (Quater)",
    "Last Conversation Notes",
    "Anniversary",
    "",
    "Recipient Email",
    "Trigger hour (0–23)"
  ];

  // Ensure enough columns exist
  var requiredCols = headers.length;
  var currentCols = sheet.getMaxColumns();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }

  // Write headers
  var headerRange = sheet.getRange(1, 1, 1, requiredCols);
  headerRange.setValues([headers]);

  // Formatting
  headerRange
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#1155cc") // Navy blue
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Freeze header row
  sheet.setFrozenRows(1);

  // Optional column widths (clean layout)
  sheet.setColumnWidths(1, requiredCols, 150);
  sheet.setColumnWidth(17, 300); // Notes column wider

  // Default values
  sheet.getRange("T2").setValue(
    Session.getEffectiveUser().getEmail()
  );
  sheet.getRange("U2").setValue(9);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Setup complete ✔ Sheet initialized and formatted."
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔔 Setup")
    .addItem("Setup sheet", "setupSheet")
    .addToUi();
}
