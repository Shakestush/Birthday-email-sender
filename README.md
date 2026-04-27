# Birthday-email-sender
# 🎂 Birthday Email Sender — KSG

A Google Apps Script that automatically sends personalised HTML birthday emails to employees based on data stored in a Google Sheet. Runs daily on a scheduled trigger and requires zero manual intervention once set up.

---

## Features

- Automatically detects employee birthdays by comparing today's date against a Google Sheet
- Sends beautifully formatted HTML birthday emails with a plain text fallback
- Includes the employee's name and designation in the personalised message
- Logs all activity — sent emails, skipped rows, and errors
- Built-in test mode that previews matches without sending real emails
- One-time trigger setup that runs daily at 7:00 AM
- Graceful error handling — one bad row does not stop the rest

---

## Prerequisites

- A Google account with access to Google Sheets and Google Apps Script
- A Google Sheet named **"Data Sheet"** with the following column structure:

| Column A | Column B | Column C | Column D |
|----------|----------|----------|----------|
| Name | Email | Birthday | Designation |

> Dates in Column C should be in a format Google Sheets recognises as a date (e.g. `15/04/1990` or `April 15, 1990`).

---

## Setup Instructions

### Step 1 — Prepare your Google Sheet

1. Open Google Sheets and create or open your employee spreadsheet.
2. Rename the relevant sheet tab to **Data Sheet**.
3. Add your column headers in Row 1: `Name`, `Email`, `Birthday`, `Designation`.
4. Populate the rows with employee data from Row 2 onwards.

### Step 2 — Open Apps Script

1. In your Google Sheet, click **Extensions** → **Apps Script**.
2. Delete any existing code in the editor.
3. Paste the full contents of `birthday_email_sender.gs` into the editor.
4. Click **Save** (or press `Ctrl+S` / `Cmd+S`).

### Step 3 — Set Up the Daily Trigger

Run the `setupDailyTrigger` function **once** to schedule automatic daily execution:

1. In the Apps Script editor, select `setupDailyTrigger` from the function dropdown.
2. Click **Run**.
3. Approve any permissions requested by Google.

This sets the script to check for birthdays every day at **7:00 AM** in your spreadsheet's timezone.

> You only need to run `setupDailyTrigger` once. Running it again removes the existing trigger and creates a fresh one to avoid duplicates.

---

## Testing

### Preview Matches Without Sending Emails

Run `testBirthdayEmails` to see which employees would receive an email today, without actually sending anything:

1. Select `testBirthdayEmails` from the function dropdown.
2. Click **Run**.
3. Open **View → Logs** to see the output.

### Send a Test Email to Yourself

To preview the actual HTML email in your inbox:

1. Open the script and find the `sendTestEmail` function.
2. Replace `"your.email@example.com"` with your own email address.
3. Select `sendTestEmail` from the dropdown and click **Run**.

---

## How It Works

1. The script reads all rows from the **Data Sheet** on each run.
2. It compares each employee's birthday (month and day only) against today's date.
3. If there is a match, it sends an HTML email to that employee's address.
4. All activity is written to the Apps Script execution log.

Year of birth is intentionally ignored — the script matches on **month and day only**, so it fires correctly every year without any updates.

---

## File Structure

```
birthday_email_sender.gs
│
├── sendBirthdayEmails()       — Main function. Reads sheet and sends emails.
├── getHTMLBirthdayMessage()   — Returns the formatted HTML email body.
├── getPlainTextBirthdayMessage() — Returns plain text fallback for the email.
├── escapeHTML()               — Sanitises name and designation to prevent HTML injection.
├── testBirthdayEmails()       — Dry-run mode. Logs matches without sending.
├── sendTestEmail()            — Sends a sample email to a specified address.
└── setupDailyTrigger()        — Creates a daily time-based trigger at 7:00 AM.
```

---

## Logging

All runs produce log output viewable under **View → Logs** in Apps Script:

| Symbol | Meaning |
|--------|---------|
| ✅ | Email sent successfully |
| ℹ️ | Employee found but birthday is not today |
| ⚠️ | Row skipped — missing name, email, or birthday |
| ❌ | Error encountered on a row |

A summary line is printed at the end of each run:
```
🎂 Birthday check completed. Sent: 2, Errors: 0
```

---

## Customisation

**Change the trigger time**
In `setupDailyTrigger`, update `.atHour(7)` to your preferred hour (24-hour format).

**Change the organisation name**
Search for `KSG` in the script and replace it with your organisation's name.

**Modify the email content**
Edit the `getHTMLBirthdayMessage` or `getPlainTextBirthdayMessage` functions to adjust the message, colours, or layout.

**Add CC or BCC recipients**
In `sendBirthdayEmails`, extend the `MailApp.sendEmail` call:
```javascript
MailApp.sendEmail({
  to: email,
  cc: "manager@example.com",
  subject: subject,
  body: plainTextMessage,
  htmlBody: htmlMessage
});
```

---

## Troubleshooting

**Emails are not being sent**
- Confirm the sheet tab is named exactly **Data Sheet** (case-sensitive).
- Check that the Birthday column contains valid date values, not plain text.
- Open **View → Logs** to see if any rows are being skipped with a warning.

**Trigger is not running**
- Go to **Triggers** (clock icon in the left sidebar of Apps Script) and confirm the trigger exists.
- If not, run `setupDailyTrigger` again.

**Permission errors**
- The script requires permission to send email and access your spreadsheet. Re-run any function and approve the permissions prompt when it appears.

**Dates not matching**
- Ensure the spreadsheet's timezone matches your expected timezone. You can check this in Google Sheets under **File → Settings**.

---

## Security Notes

- The `escapeHTML` function sanitises all user-supplied values (name and designation) before inserting them into the HTML email body, preventing HTML injection.
- No sensitive data is logged — only names, designations, and email addresses appear in logs.
- The script only has access to the spreadsheet it is bound to and the Gmail account that authorised it.

---

## License
 All rights reserved.
