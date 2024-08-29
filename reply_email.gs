/**
 * Sends an email reply to the last email for each unique UUID in the sheet.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 */
function sendReplyToLastEmailForEachUUID(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  // Option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine) {
    subjectLine = Browser.inputBox("Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
      "draft message you would like to use:",
      Browser.Buttons.OK_CANCEL);

    if (subjectLine === "cancel" || subjectLine == "") {
      return; // If no subject line, finishes up
    }
  }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains column headings
  const heads = data.shift();

  // Gets the index of the UUID column
  const uuidColIdx = heads.indexOf(UUID_COL);
  var replyColIdx = heads.indexOf(REPLY_TIME_COL);
  // If reply time or text columns don't exist, add them
  if (replyColIdx === -1) {
    sheet.getRange(1, heads.length + 1).setValue(REPLY_TIME_COL);
    replyColIdx = heads.length;
  }

  if (uuidColIdx === -1) {
    throw new Error("UUID column not found");
  }

  // Converts 2D array into an object array
  const obj = data.map(r => heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}));

  // Creates an array to record sent emails
  const out = [];

  // Loop through each unique UUID
  obj.forEach(row => {
    if (row[UUID_COL] == '') {
      out.push('')
    } else {
      const thread = searchEmailsByUUID(row[UUID_COL])[0]; // Call function from another file

      if (thread) { // Ensure we only reply to emails in the inbox
        if (row[REPLY_TIME_COL] == '') {

          const threadId = thread.getId();
          const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
          const hiddenUuid = '<div style="color:white;font-size:1px;">UUID: ' + row[UUID_COL] + '</div>';

          var draft = thread.createDraftReply("");

          draft.update(row[RECIPIENT_COL], subjectLine, msgObj.text, {
            htmlBody: msgObj.html + hiddenUuid,
            attachments: msgObj.attachments, inlinImages: msgObj.inlineImages
          })

          var sent_msg = draft.send()

          // Edits cell to record email sent date
          Logger.log(sent_msg.getDate())
          out.push([sent_msg.getDate()]);
        }
        else {
          out.push([row[REPLY_TIME_COL]]);
        }
      }
    }
  }
  )

  // Updates the sheet with new data
  sheet.getRange(2, replyColIdx + 1, out.length).setValues(out);
}

function searchEmailsByUUID(uuid) {
  var query = 'uuid: "' + uuid + '"';
  return GmailApp.search(query);
}
