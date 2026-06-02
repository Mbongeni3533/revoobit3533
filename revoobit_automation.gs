// ============================================================
// REVOOBIT LEAD AUTOMATION — Mbongeni Shongwe
// ============================================================
// HOW TO INSTALL:
// 1. Open your Google Sheet (Revoobit Leads CRM)
// 2. Click Extensions → Apps Script
// 3. Delete any existing code in the editor
// 4. Paste this entire file
// 5. Click Save (floppy disk icon)
// 6. Set Trigger 1: sendWelcomeEmail → On form submit
// 7. Set Trigger 2: checkFollowUps → Time-driven → Daily → 9am
//
// HOW TO SET TRIGGERS:
// In Apps Script editor → Click the clock icon (Triggers)
// → Add Trigger → Choose function → Choose event type
// ============================================================


// ============================================================
// TRIGGER 1: Runs automatically every time someone submits
// the Google Form. Sends them a personalised welcome email
// and marks their status as "Contacted" in your Sheet.
// ============================================================

function sendWelcomeEmail(e) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Form Responses 1"); // Change if your sheet tab has a different name

  var row    = e.values;           // All form field values from the submission
  var name   = row[1];             // Column B — Full Name
  var phone  = row[2];             // Column C — WhatsApp Number
  var email  = row[3];             // Column D — Email Address
  var interest = row[4];           // Column E — Interest (Health / Income / Both)
  var source   = row[5] || "N/A";  // Column F — How they heard about you

  // ── Build the personalised email body based on their interest ──
  var healthLink  = "https://mbongeni3533.github.io/revoobit3533/";
  var incomeLink  = "https://mbongeni3533.github.io/revoobit3533/#join";
  var waLink      = "https://wa.me/26878123067";
  var telegramLink= "https://t.me/mbongeni3533";

  var relevantLink = "";
  var relevantText = "";

  if (interest.toLowerCase().includes("health")) {
    relevantLink = healthLink;
    relevantText = "Watch the full Miira-Cell+ product presentation here:\n" + healthLink;
  } else if (interest.toLowerCase().includes("income") || interest.toLowerCase().includes("passive")) {
    relevantLink = incomeLink;
    relevantText = "See how the Revoobit binary income plan works here:\n" + incomeLink;
  } else {
    // Both health and income
    relevantLink = healthLink;
    relevantText = "Watch the full product presentation (health + income overview):\n" + healthLink;
  }

  var subject = "Hi " + name + " — your Revoobit information pack from Mbongeni";

  var body =
    "Hi " + name + "!\n\n" +
    "Thank you for reaching out — I'm Mbongeni Shongwe, Revoobit affiliate " +
    "from Eswatini, and I'm personally responding to your request.\n\n" +
    "Based on your interest in " + interest + ", here is your next step:\n\n" +
    relevantText + "\n\n" +
    "─────────────────────────────\n" +
    "WHAT REVOOBIT OFFERS YOU:\n" +
    "─────────────────────────────\n" +
    "🌿 HEALTH: Miira-Cell+ — a plant stem cell supplement with 13 natural " +
    "ingredients supporting 400+ health conditions including diabetes, blood " +
    "pressure, arthritis, skin problems, and more.\n\n" +
    "💰 INCOME: A binary compensation plan where you earn through your left " +
    "and right teams. Free to register. Works from your phone anywhere in Africa.\n\n" +
    "─────────────────────────────\n\n" +
    "I will reach out to you on WhatsApp within 24 hours at the number you " +
    "provided. If you'd like to connect sooner:\n\n" +
    "💬 WhatsApp me: " + waLink + "\n" +
    "✈  Telegram: " + telegramLink + "\n\n" +
    "Talk soon,\n" +
    "Mbongeni Shongwe\n" +
    "Revoobit ID: Mbongeni3533\n" +
    "Eswatini, Africa\n" +
    "─────────────────────────────\n" +
    "This email was sent because you submitted a request at " +
    "https://mbongeni3533.github.io/revoobit3533";

  // ── Send the email ──
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body
  });

  // ── Update the Sheet: add Status and Follow-up Date columns ──
  var lastRow = sheet.getLastRow();

  // Column G — Status
  sheet.getRange(lastRow, 7).setValue("Contacted");

  // Column H — Follow-up date (3 days from now)
  var followUpDate = new Date();
  followUpDate.setDate(followUpDate.getDate() + 3);
  sheet.getRange(lastRow, 8).setValue(followUpDate);

  // Column I — Source (how they found you)
  sheet.getRange(lastRow, 9).setValue(source);

  Logger.log("Welcome email sent to: " + name + " (" + email + ")");
}


// ============================================================
// TRIGGER 2: Runs daily at 9am. Checks your Sheet for any
// leads whose follow-up date is TODAY and sends you a
// reminder email so you never miss a follow-up.
// ============================================================

function checkFollowUps() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Form Responses 1");

  var data     = sheet.getDataRange().getValues();
  var today    = new Date();
  today.setHours(0, 0, 0, 0);

  var dueToday = [];

  for (var i = 1; i < data.length; i++) {
    var status     = data[i][6] || "";  // Column G — Status
    var followDate = data[i][7];        // Column H — Follow-up Date

    // Skip leads who have already joined or are marked cold
    if (status === "Joined" || status === "Cold") continue;

    if (followDate) {
      var fDate = new Date(followDate);
      fDate.setHours(0, 0, 0, 0);

      if (fDate.getTime() === today.getTime()) {
        dueToday.push({
          name:     data[i][1],  // Column B — Name
          phone:    data[i][2],  // Column C — WhatsApp
          email:    data[i][3],  // Column D — Email
          interest: data[i][4],  // Column E — Interest
          status:   status
        });
      }
    }
  }

  // ── Send reminder only if there are leads due today ──
  if (dueToday.length === 0) {
    Logger.log("No follow-ups due today.");
    return;
  }

  var myEmail  = Session.getActiveUser().getEmail();
  var subject  = "🔔 Revoobit Follow-ups Due Today — " + dueToday.length + " lead(s)";

  var body = "Good morning Mbongeni!\n\n" +
    "You have " + dueToday.length + " Revoobit lead(s) to follow up with today:\n\n" +
    "─────────────────────────────\n";

  dueToday.forEach(function(lead, idx) {
    body +=
      (idx + 1) + ". " + lead.name + "\n" +
      "   WhatsApp: " + lead.phone + "\n" +
      "   Email: " + lead.email + "\n" +
      "   Interest: " + lead.interest + "\n" +
      "   Status: " + lead.status + "\n" +
      "   Quick message: wa.me/" + lead.phone.replace(/\D/g,"") + "\n\n";
  });

  body +=
    "─────────────────────────────\n" +
    "Suggested follow-up message (Day 3):\n\n" +
    '"Hi [Name]! Following up on what we chatted about. I wanted to share ' +
    'one thing: 90% of chronic health problems start with damaged cells. ' +
    'Miira-Cell+ addresses this at the root. Would you like to watch a ' +
    '5-minute product overview? ' + "https://mbongeni3533.github.io/revoobit3533/" + '"\n\n' +
    "Open your Revoobit Leads CRM sheet to update their status after contact:\n" +
    "https://docs.google.com/spreadsheets/\n\n" +
    "— Your Revoobit Automation";

  MailApp.sendEmail({
    to: myEmail,
    subject: subject,
    body: body
  });

  Logger.log("Follow-up reminder sent for " + dueToday.length + " lead(s).");
}


// ============================================================
// OPTIONAL: Run this once manually to set up column headers
// in your Sheet if they are not already there.
// In Apps Script: click Run → setupSheetHeaders
// ============================================================

function setupSheetHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Form Responses 1");

  var headers = [
    "Timestamp",       // A — auto from form
    "Full Name",       // B — auto from form
    "WhatsApp Number", // C — auto from form
    "Email Address",   // D — auto from form
    "Interest",        // E — auto from form
    "Source",          // F — auto from form
    "Status",          // G — set by script (New/Contacted/Warm/Joined/Cold)
    "Follow-up Date",  // H — set by script
    "Notes"            // I — fill in manually after each call/chat
  ];

  // Only add headers to row 1 if not already set
  var firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === "Timestamp" || firstCell === "") {
    sheet.getRange(1, 7).setValue("Status");
    sheet.getRange(1, 8).setValue("Follow-up Date");
    sheet.getRange(1, 9).setValue("Notes");

    // Bold and colour the header row
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold");
    sheet.getRange(1, 7, 1, 3).setBackground("#fce8b2");
  }

  Logger.log("Headers set up successfully.");
}
