/**
 * JCF HR Follow-Up Automation — Google Apps Script
 * =================================================
 * Paste this into: Google Sheet > Extensions > Apps Script
 * Then set a Time-Driven Trigger to run sendFollowUps() daily.
 *
 * Sheet columns expected (row 1 = headers):
 *   A: Name | B: Email | C: Role Applied | D: Status
 *   E: Last Contact Date | F: Follow-Up Sent | G: Notes
 */

var SHEET_NAME    = "Candidates";
var HR_NAME       = "Disha";
var ORG_NAME      = "JCF";
var SKIP_STATUSES = ["Hired", "Rejected", "Withdrew", "On Hold"];

// Days to wait before follow-up per status
var FOLLOWUP_DAYS = {
  "Applied":             2,
  "Shortlisted":         3,
  "Interview Scheduled": 1,
  "Under Review":        4
};

// -----------------------------------------------------------------------
// MAIN: called by the daily trigger
// -----------------------------------------------------------------------
function sendFollowUps() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var headers = data[0];
  var COL = {};
  headers.forEach(function(h, i) { COL[h.trim()] = i; });

  var sentCount = 0;

  for (var r = 1; r < data.length; r++) {
    var row    = data[r];
    var name   = row[COL["Name"]];
    var email  = row[COL["Email"]];
    var role   = row[COL["Role Applied"]] || "the open position";
    var status = row[COL["Status"]] || "";
    var lastContact  = row[COL["Last Contact Date"]];
    var followupSent = row[COL["Follow-Up Sent"]];

    // Skip terminal statuses or missing email
    if (!email || SKIP_STATUSES.indexOf(status) !== -1) continue;

    // Skip if already sent today
    if (followupSent instanceof Date) {
      var sentDate = new Date(followupSent);
      sentDate.setHours(0, 0, 0, 0);
      if (sentDate.getTime() === today.getTime()) continue;
    }

    // Check elapsed days
    var gapDays = FOLLOWUP_DAYS[status] || 3;
    if (lastContact instanceof Date) {
      var daysSince = Math.floor((today - lastContact) / 86400000);
      if (daysSince < gapDays) continue;
    }

    // Build and send email
    var firstName = name.toString().split(" ")[0];
    var mail = buildEmailContent(status, firstName, role);

    try {
      GmailApp.sendEmail(email, mail.subject, "", { htmlBody: mail.body, name: HR_NAME + " | " + ORG_NAME });

      // Stamp Follow-Up Sent and update Last Contact Date
      sheet.getRange(r + 1, COL["Follow-Up Sent"] + 1).setValue(new Date());
      sheet.getRange(r + 1, COL["Last Contact Date"] + 1).setValue(new Date());
      // Highlight row green
      sheet.getRange(r + 1, 1, 1, headers.length)
           .setBackground("#D9F0D3");

      sentCount++;
      Logger.log("✓ Sent to " + name + " <" + email + ">");
    } catch(e) {
      Logger.log("✗ Failed for " + name + ": " + e.message);
    }
  }

  Logger.log("=== Done: " + sentCount + " follow-up(s) sent ===");
  if (sentCount > 0) {
    sendSummaryToHR(sentCount);
  }
}

// -----------------------------------------------------------------------
// EMAIL TEMPLATES
// -----------------------------------------------------------------------
function buildEmailContent(status, firstName, role) {
  var templates = {
    "Applied": {
      subject: "Update on Your Application – " + role + " at " + ORG_NAME,
      body: "Dear " + firstName + ",<br><br>" +
        "Thank you for applying for the <b>" + role + "</b> position at " + ORG_NAME + ". " +
        "We have received your application and our team is currently reviewing it.<br><br>" +
        "We aim to get back to you within 3–5 business days. Feel free to reach out with any questions.<br><br>" +
        "Warm regards,<br><b>" + HR_NAME + "</b><br>HR Team, " + ORG_NAME
    },
    "Shortlisted": {
      subject: "Great News! You've Been Shortlisted – " + role + " at " + ORG_NAME,
      body: "Dear " + firstName + ",<br><br>" +
        "We are pleased to inform you that you have been <b>shortlisted</b> for the <b>" + role + "</b> role at " + ORG_NAME + ". 🎉<br><br>" +
        "Our team will be in touch shortly to schedule the next steps.<br><br>" +
        "Best regards,<br><b>" + HR_NAME + "</b><br>HR Team, " + ORG_NAME
    },
    "Interview Scheduled": {
      subject: "Interview Reminder – " + role + " at " + ORG_NAME,
      body: "Dear " + firstName + ",<br><br>" +
        "This is a friendly reminder about your upcoming interview for the <b>" + role + "</b> position at " + ORG_NAME + ".<br><br>" +
        "Please confirm your availability by replying to this email.<br><br>" +
        "Looking forward to speaking with you!<br><br>" +
        "<b>" + HR_NAME + "</b><br>HR Team, " + ORG_NAME
    },
    "Under Review": {
      subject: "Application Status Update – " + role + " at " + ORG_NAME,
      body: "Dear " + firstName + ",<br><br>" +
        "We wanted to let you know that your application for <b>" + role + "</b> at " + ORG_NAME + " is still under review.<br><br>" +
        "We appreciate your patience and will be in touch soon.<br><br>" +
        "<b>" + HR_NAME + "</b><br>HR Team, " + ORG_NAME
    }
  };

  return templates[status] || {
    subject: "Follow-Up: Your Application at " + ORG_NAME,
    body: "Dear " + firstName + ",<br><br>We are following up on your application for <b>" + role + "</b> at " + ORG_NAME + ".<br><br>" +
      "Please reply if you have any questions.<br><br><b>" + HR_NAME + "</b><br>HR Team, " + ORG_NAME
  };
}

// -----------------------------------------------------------------------
// SUMMARY EMAIL TO HR
// -----------------------------------------------------------------------
function sendSummaryToHR(count) {
  var hrEmail = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(hrEmail,
    "[JCF HR Bot] " + count + " Follow-Up(s) Sent Today",
    "",
    { htmlBody: "<p>The automated follow-up script ran successfully.<br>" +
      "<b>" + count + "</b> candidate follow-up email(s) were sent on " + new Date().toDateString() + ".</p>" +
      "<p>Check the <a href='" + SpreadsheetApp.getActiveSpreadsheet().getUrl() + "'>Candidate Tracker</a> for details.</p>" }
  );
}

// -----------------------------------------------------------------------
// SETUP: Add time-driven trigger programmatically
// -----------------------------------------------------------------------
function createDailyTrigger() {
  // Delete existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendFollowUps")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  Logger.log("Daily trigger set for 09:00 AM");
}
