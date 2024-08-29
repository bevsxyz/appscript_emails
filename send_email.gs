/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine) {
    subjectLine = Browser.inputBox("Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
      "draft message you would like to mail merge with:",
      Browser.Buttons.OK_CANCEL);

    if (subjectLine === "cancel" || subjectLine == "") {
      // If no subject line, finishes up
      return;
    }
  }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift();

  var emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  var uuidColIdx = heads.indexOf(UUID_COL);

  // If email sent or uuid columns don't exist, add them
  if (emailSentColIdx === -1) {
    sheet.getRange(1, heads.length + 1).setValue(EMAIL_SENT_COL);
    emailSentColIdx = heads.length;
  }
  if (uuidColIdx === -1) {
    sheet.getRange(1, heads.length + 1).setValue(UUID_COL);
    uuidColIdx = heads.length;
  }

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];
  const uniqueID = [];

  // Loops through all the rows of data
  obj.forEach(function (row, rowIdx) {
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[EMAIL_SENT_COL] == '') {
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        var uuid = Utilities.getUuid();
        var hiddenUuid = '<div style="color:white;font-size:1px;">UUID: ' + uuid + '</div>';

        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        MailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html + hiddenUuid,
          // bcc: row[RECIPIENT_COL],
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Edits cell to record email sent date and uuid of the email
        out.push([new Date()]);
        uniqueID.push([uuid])
      } catch (e) {
        // modify cell to record error
        out.push([e.message]);
        uniqueID.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
      uniqueID.push([row[UUID_COL]]);
    }
  });

  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
  sheet.getRange(2, uuidColIdx + 1, uniqueID.length).setValues(uniqueID);
}