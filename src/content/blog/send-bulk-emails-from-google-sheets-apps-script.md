---
title: 'How to Send Personalized Bulk Emails from Google Sheets (Free Apps Script)'
description: 'Send personalized emails to a list in Google Sheets with a free Apps Script. Copy-paste code, a status column so nobody gets emailed twice, and daily limits explained.'
pubDate: '2026-10-22'
tags: ['Google Apps Script', 'Google Sheets', 'Email Automation']
# Remove this line (or set it to false) when you have reviewed the post and want it live.
draft: true
---

Copy-pasting the same email 50 times and changing the name each time is exactly the kind of boring
task a computer should do. In this experiment I send personalized emails to everyone in a Google
Sheet with about 25 lines of free Google Apps Script. No add-ons, no paid mail-merge tools.

## What you need

- A Google account (free Gmail works)
- A Google Sheet with your contacts

## Step 1: Set up the sheet

Create a sheet tab named **Contacts** with these column headers in row 1:

| Name  | Email             | Status |
| ----- | ----------------- | ------ |
| Priya | priya@example.com |        |
| Rahul | rahul@example.com |        |

Leave **Status** empty. The script writes `Sent` there, so if you run it again, nobody gets the
same email twice.

## Step 2: Add the script

In your sheet, open **Extensions → Apps Script**, delete the sample code and paste this:

```js
const SUBJECT = 'Quick question about your website';

function sendBulkEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Contacts');
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const col = (name) => headers.indexOf(name);

  let remaining = MailApp.getRemainingDailyQuota();

  rows.forEach((row, i) => {
    if (row[col('Status')] === 'Sent' || !row[col('Email')] || remaining <= 0) return;

    const body =
      `Hi ${row[col('Name')]},\n\n` +
      `I noticed a couple of quick SEO wins for your site and wanted to share them.\n\n` +
      `Cheers,\nSupriyo`;

    MailApp.sendEmail(row[col('Email')], SUBJECT, body);
    sheet.getRange(i + 2, col('Status') + 1).setValue('Sent');
    remaining--;
  });
}

// Adds a "Bulk Email" menu to the sheet so you can run it without opening the editor.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bulk Email').addItem('Send emails', 'sendBulkEmails').addToUi();
}
```

Save with **Ctrl + S**.

## Step 3: Run it

1. Select `sendBulkEmails` in the toolbar and click **Run**.
2. The first time, Google asks for permission to send email as you. Review and allow it.
3. Check the sheet: every row that was emailed now says **Sent**.

Next time, just reload the sheet and use **Bulk Email → Send emails**.

## How many emails can you send?

Google limits how many recipients Apps Script can email per day: roughly **100 per day on a free
Gmail account** and **1,500 on Google Workspace**. The script checks the remaining quota with
`MailApp.getRemainingDailyQuota()` and stops before hitting it. Unsent rows stay empty, so you can
run it again tomorrow and it continues where it left off.

## Ideas to take it further

- Add more columns (company, website) and use them in the email body.
- Send HTML emails with the `htmlBody` option.
- Run it automatically every morning with a time-driven trigger.

<!-- TODO: Add your own result here, e.g. how long this took before vs after, or a screenshot of the sheet. First-hand results are what make a post rank. -->
