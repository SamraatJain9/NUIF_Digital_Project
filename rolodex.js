/*
  Authors: Samraat Jain, James Delin, Sarah Rafiepour, Ryan Duong, Shalom Ademuwagun
  Version: // A Rolodex Script v1.6 — 2026-02-23
  This script:
  - Automatically sets up the sheet with the necessary headers and formatting for A Rolodex Script.
  - Reads your contact sheet.
  - Sends a daily reminder email to you (or to the address in T2).
  - Never sends data anywhere else.
  - Requires permission to: read your sheet, send email, and create triggers.
  - Your memory assistant for 1,000+ connections.
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
  applyDateFieldFormatting(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast(
      "Setup complete ✔ Sheet initialized and formatted."
  );
}

function applyDateFieldFormatting(sheet) {
  var dateRanges = ["L2:L", "N2:N", "O2:O", "R2:R"];
  var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText("Enter a valid date (yyyy/MM/dd)")
      .build();

  dateRanges.forEach(function (a1) {
    var range = sheet.getRange(a1);
    range.setNumberFormat("yyyy/MM/dd");
    range.setDataValidation(dateRule);
  });
}
function sendReminders(batchStart) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var recipient =
      sheet.getRange("T2").getValue() ||
      Session.getEffectiveUser().getEmail();

  // Use spreadsheet timezone as single source of truth
  var TZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  // "Today" as a pure calendar date in spreadsheet timezone
  var today = new Date(
      Utilities.formatDate(new Date(), TZ, "yyyy/MM/dd")
  );

  var nameCol = headers.indexOf("Name");
  var emailCol = headers.indexOf("Email");
  var phoneCol = headers.indexOf("Phone Number");
  var timezoneCol = headers.indexOf("Timezone");
  var companyCol = headers.indexOf("Company");
  var titleCol = headers.indexOf("Title");
  var notesCol = headers.indexOf("Last Conversation Notes");
  var birthdayCol = headers.indexOf("Birthday");
  var anniversaryCol = headers.indexOf("Anniversary");
  var lastInteractionCol = headers.indexOf("Last Interaction");
  var touchIntervalCol = headers.indexOf("Touch Interval (Quater)");

  var rows = [];
  var batchSize = 900;
  var start = batchStart ? parseInt(batchStart, 10) : 1;
  var end = Math.min(start + batchSize, data.length);

  function escapeHtml(text) {
    if (!text) return "";
    return text
        .toString()
        .trim()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
  }

  function parseSheetDate(value) {
    if (!value) return null;
    if (value instanceof Date) {
      return new Date(
          Utilities.formatDate(value, TZ, "yyyy/MM/dd")
      );
    }
    return null;
  }

  function isSameMonthDay(d1, d2) {
    if (!d1 || !d2) return false;
    return (
        d1.getMonth() === d2.getMonth() &&
        d1.getDate() === d2.getDate()
    );
  }

  function isSameDate(d1, d2) {
    if (!d1 || !d2) return false;
    return (
        d1.getFullYear() === d2.getFullYear() &&
        d1.getMonth() === d2.getMonth() &&
        d1.getDate() === d2.getDate()
    );
  }

  function addMonths(date, months) {
    var d = new Date(date);
    var day = d.getDate();
    d.setMonth(d.getMonth() + months);

    // Handle month rollover (e.g., Jan 31 → Feb)
    if (d.getDate() < day) {
      d.setDate(0);
    }
    return d;
  }

  if (start === 1) {
    rows.push(
        "<tr style='background:#f2f2f2;'>" +
        "<th>Name</th>" +
        "<th>Trigger Type</th>" +
        "<th>Email / Phone</th>" +
        "<th>Time Zone</th>" +
        "<th>Company & Title</th>" +
        "<th>Last Conversation Notes</th>" +
        "</tr>"
    );
  }

  for (var i = start; i < end; i++) {
    var row = data[i];

    var name = escapeHtml(row[nameCol]);
    var email = escapeHtml(row[emailCol]);
    var phone = escapeHtml(row[phoneCol]);
    var company = escapeHtml(row[companyCol]);
    var title = escapeHtml(row[titleCol]);
    var tzCell = escapeHtml(row[timezoneCol]);
    var notes = escapeHtml(row[notesCol]);

    var birthday = parseSheetDate(row[birthdayCol]);
    var anniversary = parseSheetDate(row[anniversaryCol]);
    var lastInteraction = parseSheetDate(row[lastInteractionCol]);
    var touchIntervalQuarter = parseInt(row[touchIntervalCol], 10);

    var triggers = [];

    if (!name && !email && !phone) continue;

    if (isSameMonthDay(today, birthday)) {
      triggers.push("Birthday");
    }

    if (isSameMonthDay(today, anniversary)) {
      triggers.push("Anniversary");
    }

    if (lastInteraction && touchIntervalQuarter) {
      var nextTouch = addMonths(
          lastInteraction,
          touchIntervalQuarter * 3
      );
      if (isSameDate(today, nextTouch)) {
        triggers.push("Touch Interval");
      }
    }

    if (triggers.length) {
      rows.push(
          "<tr>" +
          "<td>" + name + "</td>" +
          "<td>" + triggers.join(", ") + "</td>" +
          "<td>" + email + (phone ? "<br>" + phone : "") + "</td>" +
          "<td>" + tzCell + "</td>" +
          "<td>" + company + (title ? " — " + title : "") + "</td>" +
          "<td>" + notes + "</td>" +
          "</tr>"
      );
    }
  }

  var cache = CacheService.getScriptCache();
  var prev = cache.get("reminderRows");
  var accumulated = prev ? JSON.parse(prev) : [];
  accumulated = accumulated.concat(rows);
  cache.put("reminderRows", JSON.stringify(accumulated), 3600);

  if (end < data.length) {
    ScriptApp.newTrigger("continueReminders")
        .timeBased()
        .after(2 * 60 * 1000)
        .create();
    PropertiesService.getScriptProperties().setProperty("nextStart", end);
  } else {
    var allRows = JSON.parse(cache.get("reminderRows") || "[]");
    cache.remove("reminderRows");

    if (allRows.length > 1) {
      var htmlBody =
          "<html><body>" +
          "<table border='1' cellpadding='5' cellspacing='0' " +
          "style='border-collapse:collapse;width:100%;'>" +
          allRows.join("") +
          "</table></body></html>";

      MailApp.sendEmail({
        to: recipient,
        subject: "Rolodex Reminder Notification",
        htmlBody: htmlBody
      });
    }

    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === "continueReminders") {
        ScriptApp.deleteTrigger(t);
      }
    });

    PropertiesService.getScriptProperties().deleteProperty("nextStart");
  }
}

function continueReminders() {
  var nextStart =
      PropertiesService.getScriptProperties().getProperty("nextStart");
  sendReminders(nextStart);
}

function setupDailyTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var hour = parseInt(sheet.getRange("U2").getDisplayValue().trim(), 10);
  if (isNaN(hour) || hour < 0 || hour > 23) hour = 9;

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sendReminders") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("sendReminders")
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .inTimezone(SpreadsheetApp.getActive().getSpreadsheetTimeZone())
      .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
      "Daily reminder set for " +
      hour +
      ":00 (" +
      SpreadsheetApp.getActive().getSpreadsheetTimeZone() +
      ")"
  );
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("All triggers removed.");
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu("🔔 Setup")
      .addItem("Setup sheet", "setupSheet")
      .addToUi();

  SpreadsheetApp.getUi()
      .createMenu("🔔 Reminders")
      .addItem("Run reminders now", "sendReminders")
      .addItem("Set up daily trigger", "setupDailyTrigger")
      .addSeparator()
      .addItem("Remove all triggers", "removeAllTriggers")
      .addToUi();
}
