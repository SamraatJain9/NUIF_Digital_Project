function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var headers = [
    "Name","Email","Phone Number","LinkedIn","Company","Title","Industry",
    "Country of Residence","Religion","Birthday","Holidays",
    "Last Meeting","Contact Interval","Anniversary","",
    "Recipient Email","Trigger hour (0–23)"
  ];

  var requiredCols = headers.length;
  var currentCols = sheet.getMaxColumns();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }

  var headerRange = sheet.getRange(1,1,1,requiredCols);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold")
             .setFontColor("#FFFFFF")
             .setBackground("#1155cc")
             .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, requiredCols, 150);

  sheet.getRange("J2:J").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("L2:L").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("N2:N").setNumberFormat("yyyy-mm-dd");

  var options = [
    "1 week","2 weeks","3 weeks",
    "1 month","2 months","3 months",
    "6 months","12 months"
  ];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("M2:M").setDataValidation(rule);

  sheet.getRange("P2").setValue(Session.getEffectiveUser().getEmail());
  sheet.getRange("Q2").setValue(9);

  SpreadsheetApp.getActiveSpreadsheet().toast("Setup complete. Date columns now have calendar picker.");
}

function parseDate(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") return value;
  var p = value.toString().split("-");
  if (p.length !== 3) return null;
  return new Date(p[0], p[1]-1, p[2]);
}

function intervalToDays(text) {
  if (!text) return 0;
  if (text == "1 week") return 7;
  if (text == "2 weeks") return 14;
  if (text == "3 weeks") return 21;
  if (text == "1 month") return 30;
  if (text == "2 months") return 60;
  if (text == "3 months") return 90;
  if (text == "6 months") return 180;
  if (text == "12 months") return 365;
  return 0;
}

function sendReminders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var recipient = sheet.getRange("P2").getValue() || Session.getEffectiveUser().getEmail();

  var today = new Date();
  var todayMonth = today.getMonth();
  var todayDate = today.getDate();

  var nameCol = headers.indexOf("Name");
  var emailCol = headers.indexOf("Email");
  var phoneCol = headers.indexOf("Phone Number");
  var companyCol = headers.indexOf("Company");
  var titleCol = headers.indexOf("Title");
  var birthdayCol = headers.indexOf("Birthday");
  var anniversaryCol = headers.indexOf("Anniversary");
  var lastInteractionCol = headers.indexOf("Last Meeting");
  var intervalCol = headers.indexOf("Contact Interval");

  var rows = [];

  rows.push(
    "<tr style='background:#f2f2f2;'>" +
    "<th>Name</th>" +
    "<th>Trigger</th>" +
    "<th>Contact</th>" +
    "<th>Company</th>" +
    "</tr>"
  );

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var name = row[nameCol];
    var email = row[emailCol];
    var phone = row[phoneCol];
    var company = row[companyCol];
    var title = row[titleCol];

    var birthday = parseDate(row[birthdayCol]);
    var anniversary = parseDate(row[anniversaryCol]);
    var lastInteraction = parseDate(row[lastInteractionCol]);
    var intervalDays = intervalToDays(row[intervalCol]);

    var triggers = [];

    if (!name && !email && !phone) continue;

    if (birthday) {
      if (birthday.getMonth() === todayMonth && birthday.getDate() === todayDate) {
        triggers.push("Birthday");
      }
    }

    if (anniversary) {
      if (anniversary.getMonth() === todayMonth && anniversary.getDate() === todayDate) {
        triggers.push("Anniversary");
      }
    }

    if (lastInteraction && intervalDays) {
      var nextTouch = new Date(lastInteraction);
      nextTouch.setDate(nextTouch.getDate() + intervalDays);

      if (nextTouch.getMonth() === todayMonth && nextTouch.getDate() === todayDate) {
        triggers.push("Follow up");
      }
    }

    if (triggers.length) {
      rows.push(
        "<tr>" +
        "<td>"+(name||"")+"</td>" +
        "<td>"+triggers.join(", ")+"</td>" +
        "<td>"+(email||"")+(phone?"<br>"+phone:"")+"</td>" +
        "<td>"+(company||"")+(title?" — "+title:"")+"</td>" +
        "</tr>"
      );
    }
  }

  if (rows.length > 1) {
    var html =
      "<html><body>" +
      "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;width:100%;'>" +
      rows.join("") +
      "</table></body></html>";

    MailApp.sendEmail({
      to: recipient,
      subject: "Rolodex Reminder Notification",
      htmlBody: html
    });
  }
}

function setupDailyTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var hour = parseInt(sheet.getRange("Q2").getDisplayValue().trim(),10);
  if (isNaN(hour) || hour<0 || hour>23) hour = 9;

  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction()=="sendReminders") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("sendReminders")
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast("Reminder set for "+hour+":00");
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("Triggers removed");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Setup")
    .addItem("Setup sheet","setupSheet")
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu("Reminders")
    .addItem("Run reminders now","sendReminders")
    .addItem("Set daily reminder","setupDailyTrigger")
    .addSeparator()
    .addItem("Remove triggers","removeAllTriggers")
    .addToUi();
}