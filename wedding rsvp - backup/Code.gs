// ==========================================
// 1. TWILIO CONFIGURATION CLUSTER
// ==========================================
var TWILIO_ACCOUNT_SID = "ACba8dd61543..."; // Replace with full SID from image
var TWILIO_AUTH_TOKEN = "YOUR_SHOW_REVEALED_AUTH_TOKEN"; // Replace with your authentic token
var TWILIO_SMS_NUMBER = "+14244845819";
var TWILIO_WHATSAPP_NUMBER = "+14155238886";

// ==========================================
// 2. RECEIVE SUBMISSION & SAVE AS PENDING
// ==========================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    var title = data.title || "";
    var name = data.name || "";
    var phone = data.phone || "";
    var familySide = data.familySide || "";
    var attending = data.attending || "";
    var preference = data.preference || "";
    var defaultStatus = "Pending";

    var lastRow = sheet.getLastRow();

    // Create an auto-incrementing Guest ID (e.g., G-1001, G-1002...)
    var nextIdNumber = lastRow > 1 ? lastRow : 1;
    // Derive initials from the provided name (first and last name letters).
    var initials = "";
    if (name) {
      var parts = String(name).trim().split(/\s+/);
      if (parts.length === 1) {
        initials = parts[0].substring(0, 2).toUpperCase();
      } else {
        initials = (
          parts[0].charAt(0) + parts[parts.length - 1].charAt(0)
        ).toUpperCase();
      }
    } else {
      initials = "GU"; // fallback
    }

    // Format guestId like: "CM-26001" where 260 is country code and last two digits are a sequence
    var seq = String(nextIdNumber).padStart(2, "0");
    var guestId = initials + "-260" + seq;

    // Generate a random unique 6-character string for Column H
    var uniqueTicketCode =
      "TC-" + Math.random().toString(36).substring(2, 8).toUpperCase();

    // Scan Column D (4) for Phone duplicates and Column G (7) for Status metrics
    if (lastRow > 1) {
      var phoneRange = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
      var statusRange = sheet.getRange(2, 7, lastRow - 1, 1).getValues();

      for (var i = 0; i < phoneRange.length; i++) {
        if (String(phoneRange[i][0]).trim() === String(phone).trim()) {
          var currentStatus = statusRange[i][0] || defaultStatus;
          return ContentService.createTextOutput(
            JSON.stringify({
              result: "duplicate",
              status: currentStatus,
            }),
          ).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    sheet.appendRow([
      guestId, // A (1)
      title, // B (2)
      name, // C (3)
      phone, // D (4)
      familySide, // E (5)
      attending, // F (6)
      defaultStatus, // G (7)
      uniqueTicketCode, // H (8)
      "", // I (9) - Blank placeholder for tracking logs
      "No", // J (10) - Default attendance check-in baseline
      preference, // K (11)
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({
        result: "success",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        result: "error",
        message: error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. MONITOR SHEET FOR MANUAL APPROVAL
// ==========================================
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();

  if (range.getColumnIndex() === 7 && range.getRow() > 1) {
    var newValue = range.getValue();

    if (newValue === "Approved") {
      var row = range.getRow();

      var title = sheet.getRange(row, 2).getValue();
      var name = sheet.getRange(row, 3).getValue();
      var phone = sheet.getRange(row, 4).getValue();
      var ticketCode = sheet.getRange(row, 8).getValue();
      var preference = sheet.getRange(row, 11).getValue();

      var fullName = title + " " + name;
      // Avoid calling UrlFetchApp from a simple onEdit trigger (permission-limited).
      // Instead, mark the row as queued; a separate processor will perform the network call.
      sheet.getRange(row, 9).setValue("Queued");
    }
  }
}

// Scans the sheet for Approved rows that are not yet sent and dispatches Twilio messages.
// This function must be run as an installable trigger or manually from the script editor
// because it uses UrlFetchApp.
function processPendingApprovals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowIndex = i + 2;
    var status = String(data[i][6] || "").trim(); // column G
    var sendLog = String(data[i][8] || "").trim(); // column I
    // Only process if Approved and not already sent (starts with "Sent")
    if (status === "Approved" && !sendLog.startsWith("Sent")) {
      var title = data[i][1];
      var name = data[i][2];
      var phone = data[i][3];
      var ticketCode = data[i][7];
      var preference = data[i][10];

      var fullName = (title || "") + " " + (name || "");

      try {
        var res = sendTwilioNotification(
          fullName,
          phone,
          ticketCode,
          preference,
        );
        if (res && res.success) {
          sheet
            .getRange(rowIndex, 9)
            .setValue("Sent: " + new Date().toLocaleString());
        } else {
          sheet
            .getRange(rowIndex, 9)
            .setValue("Error: " + (res && res.error ? res.error : "Failed"));
        }
      } catch (err) {
        sheet
          .getRange(rowIndex, 9)
          .setValue(
            "Error: " + (err && err.toString ? err.toString() : "Crash"),
          );
      }
    }
  }
}

// ==========================================
// 4. TWILIO DISPATCH NATIVE INTERACTION ROUTINE
// ==========================================
function sendTwilioNotification(fullName, phone, ticketCode, preference) {
  // Returns an object: {success: boolean, error?: string, response?: object}
  var formattedToPhone = String(phone || "").trim();
  if (!formattedToPhone) return { success: false, error: "Missing phone" };
  if (!formattedToPhone.startsWith("+"))
    formattedToPhone = "+" + formattedToPhone;

  var messageBody =
    "Hello " +
    fullName +
    ",\n\nYour RSVP for the wedding has been manually Approved! 🎉\n\nYour unique entry ticket code is: " +
    ticketCode +
    "\n\nPlease present this code at the venue gate for verification. See you there!";

  var url =
    "https://api.twilio.com/2010-04-01/Accounts/" +
    TWILIO_ACCOUNT_SID +
    "/Messages.json";
  var payload = { Body: messageBody };

  if (String(preference) === "WhatsApp") {
    payload.From = "whatsapp:" + TWILIO_WHATSAPP_NUMBER;
    payload.To = "whatsapp:" + formattedToPhone;
  } else {
    payload.From = TWILIO_SMS_NUMBER;
    payload.To = formattedToPhone;
  }

  var authHeader =
    "Basic " +
    Utilities.base64Encode(TWILIO_ACCOUNT_SID + ":" + TWILIO_AUTH_TOKEN);
  var options = {
    method: "post",
    headers: { Authorization: authHeader },
    payload: payload,
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var text = response.getContentText();
  var parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    parsed = { raw: text };
  }

  if (code === 200 || code === 201) {
    return { success: true, response: parsed };
  }
  return {
    success: false,
    error: (parsed && parsed.message) || (parsed && parsed.error) || text,
  };
}
