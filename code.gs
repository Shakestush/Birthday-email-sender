/**
 * Birthday Email Sender- HTML Version
 * Sends automatic birthday wishes to employees based on Google Sheet data
 * Columns expected: Name | Email | Birthday | Designation
 */

function sendBirthdayEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Sheet");
  var data = sheet.getDataRange().getValues();

  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var today = new Date();
  var todayMMDD = Utilities.formatDate(today, tz, "MM-dd");

  var sentCount = 0;
  var errorCount = 0;

  for (var i = 1; i < data.length; i++) { // skip header row
    try {
      var name = data[i][0];
      var email = data[i][1];
      var birthday = data[i][2];
      var designation = data[i][3];

      // Skip row if name or email is missing
      if (!name || !email) {
        Logger.log("⚠️ Row " + (i+1) + ": Missing name or email - skipped");
        continue;
      }

      if (birthday) {
        var bMMDD = Utilities.formatDate(new Date(birthday), tz, "MM-dd");

        if (bMMDD === todayMMDD) {
          var subject = "🎂🎉 Happy Birthday " + name + "! 🎈🎁";

          // ========== HTML FORMATTED BIRTHDAY MESSAGE ==========
          var htmlMessage = getHTMLBirthdayMessage(name, designation);
          var plainTextMessage = getPlainTextBirthdayMessage(name, designation);

          MailApp.sendEmail({
            to: email,
            subject: subject,
            body: plainTextMessage,
            htmlBody: htmlMessage
          });
          
          sentCount++;
          Logger.log("✅ Sent to " + name + " (" + designation + ") - " + email);
        } else {
          Logger.log("ℹ️ Row " + (i+1) + ": " + name + " - birthday not today (" + bMMDD + ")");
        }
      } else {
        Logger.log("⚠️ Row " + (i+1) + ": No birthday value found for " + name);
      }
    } catch (error) {
      errorCount++;
      Logger.log("❌ Error on row " + (i+1) + ": " + error.toString());
    }
  }

  Logger.log("🎂 Birthday check completed. Sent: " + sentCount + ", Errors: " + errorCount);
}

/**
 * Generates HTML formatted birthday message
 */
function getHTMLBirthdayMessage(name, designation) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        .birthday-container {
          font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
          max-width: 600px;
          margin: 0 auto;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          border-radius: 20px;
          overflow: hidden;
          box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        }
        .birthday-header {
          background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
          padding: 30px 20px;
          text-align: center;
        }
        .birthday-header h1 {
          color: white;
          font-size: 32px;
          margin: 0;
          text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        .birthday-emoji {
          font-size: 60px;
          margin: 10px 0;
        }
        .birthday-body {
          background: white;
          padding: 30px;
          color: #333;
          line-height: 1.6;
        }
        .birthday-message {
          font-size: 16px;
          margin-bottom: 20px;
        }
        .highlight {
          color: #f5576c;
          font-weight: bold;
          font-size: 18px;
        }
        .designation-badge {
          display: inline-block;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          padding: 8px 16px;
          border-radius: 50px;
          font-size: 14px;
          margin: 15px 0;
        }
        .cake-animation {
          text-align: center;
          font-size: 50px;
          margin: 20px 0;
        }
        .signature {
          margin-top: 30px;
          padding-top: 20px;
          border-top: 2px solid #eee;
          color: #888;
          font-style: italic;
        }
        .footer {
          background: #f8f9fa;
          padding: 15px;
          text-align: center;
          font-size: 12px;
          color: #888;
        }
        .confetti {
          text-align: center;
          font-size: 24px;
        }
        @media (max-width: 600px) {
          .birthday-body {
            padding: 20px;
          }
          .birthday-header h1 {
            font-size: 24px;
          }
        }
      </style>
    </head>
    <body style="margin: 0; padding: 20px; background: #f0f2f5;">
      <div class="birthday-container">
        <div class="birthday-header">
          <div class="birthday-emoji">🎂🎉🎈</div>
          <h1>Happy Birthday, ${escapeHTML(name)}!</h1>
          <div class="birthday-emoji">🎁🎊✨</div>
        </div>
        
        <div class="birthday-body">
          <div class="confetti">🎉 🎊 🎉 🎊 🎉</div>
          
          <div class="birthday-message">
            <p><strong>Dear ${escapeHTML(name)},</strong></p>
            
            <p>🎵 <em>Happy Birthday to YOU!</em> 🎵</p>
            
            <p>On your special day, we want you to know how much we appreciate having you on the team.</p>
                      
            <p>As our <span class="highlight">${escapeHTML(designation)}</span>, you bring energy, expertise, and heart to everything you do — and that makes a <strong>real difference</strong> to our team and our mission.</p>
            
            <div class="cake-animation">
              🎂 🕯️ 🎂 🕯️ 🎂
            </div>
            
            <p><strong>Today, we celebrate YOU.</strong> We hope your day is filled with:</p>
            <ul>
              <li>🎁 Joy and laughter</li>
              <li>🍰 Delicious cake (or your favorite treat!)</li>
              <li>😊 Warm smiles from loved ones</li>
              <li>✨ Everything that makes you happy</li>
            </ul>
            
            <p>Thank you for being such a wonderful part of the <strong>family</strong>. Your work doesn't go unnoticed, and today we want you to feel truly celebrated.</p>
            
            <div class="cake-animation">
              🎈 🎀 🎈 🎀 🎈
            </div>
          </div>
          
          <div class="signature">
            <strong>Warmest regards,</strong><br>
            🎂 <strong>The <institution name> Team</strong> 🎂
          </div>
        </div>
        
        <div class="footer">
          💙 You make <institution name> a great place to work! 💙
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Generates plain text version as backup for email clients that don't support HTML
 */
function getPlainTextBirthdayMessage(name, designation) {
  return `
🎂🎉 HAPPY BIRTHDAY ${name.toUpperCase()}! 🎉🎂

Dear ${name},

Happy Birthday to YOU! 🎉🎂

On your special day, we want you to know how much we appreciate having you on the team. As our ${designation}, you bring energy, expertise, and heart to everything you do — and that makes a real difference.

Today, we celebrate you. We hope your day is filled with joy, laughter, cake, and everything that makes you smile. 😊

Thank you for being part of the family.

Warmest regards,
The KSG Team

---
💙 You make <institution name> a great place to work! 💙
  `;
}

/**
 * Helper function to prevent HTML injection
 */
function escapeHTML(str) {
  if (!str) return "";
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Optional: Test function - no real emails sent
 */
function testBirthdayEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Sheet");
  var data = sheet.getDataRange().getValues();

  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var today = new Date();
  var todayMMDD = Utilities.formatDate(today, tz, "MM-dd");

  Logger.log("========== TEST MODE - NO EMAILS SENT ==========");
  Logger.log("Today's date: " + todayMMDD);
  Logger.log("");

  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    var email = data[i][1];
    var birthday = data[i][2];
    var designation = data[i][3];

    if (birthday) {
      var bMMDD = Utilities.formatDate(new Date(birthday), tz, "MM-dd");

      if (bMMDD === todayMMDD) {
        Logger.log("🎂 WOULD SEND to: " + name + " (" + designation + ") - " + email);
        Logger.log("   HTML preview: " + getHTMLBirthdayMessage(name, designation).substring(0, 200) + "...");
        Logger.log("");
      }
    }
  }

  Logger.log("========== TEST COMPLETE ==========");
}

/**
 * Setup daily trigger at 9:00 AM
 * Run this function ONCE to schedule automated birthday emails
 */
function setupDailyTrigger() {
  // Remove existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "sendBirthdayEmails") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new daily trigger
  ScriptApp.newTrigger("sendBirthdayEmails")
    .timeBased()
    .atHour(7)
    .nearMinute(0)
    .everyDays(1)
    .create();

  Logger.log("✅ Daily trigger set for 7:00 AM");
}

/**
 * Optional: Send a test email to yourself
 * Replace "your.email@example.com" with your actual email
 */
function sendTestEmail() {
  var testEmail = "your.email@example.com"; // CHANGE THIS
  var testName = "Test User";
  var testDesignation = "Test Designation";
  
  MailApp.sendEmail({
    to: testEmail,
    subject: "🎂 TEST - Birthday Email Preview",
    body: getPlainTextBirthdayMessage(testName, testDesignation),
    htmlBody: getHTMLBirthdayMessage(testName, testDesignation)
  });
  
  Logger.log("✅ Test email sent to " + testEmail);
}
