/*
  Author(s): Samraat Jain, 
  Version: // A Rolodex Script v1.4 — 2026-02-10
  This script:
  - Reads your contact sheet.
  - Sends a daily reminder email to you (or to the address in T2).
  - Never sends data anywhere else.
  - Requires permission to: read your sheet, send email, and create triggers.
  - Your memory assistant for 1,000+ connections.
*/

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
    .createMenu("🔔 Reminders")
    .addItem("Run reminders now", "sendReminders")
    .addItem("Set up daily trigger", "setupDailyTrigger")
    .addSeparator()
    .addItem("Remove all triggers", "removeAllTriggers")
    .addToUi();
}
