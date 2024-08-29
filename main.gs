/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";
const REPLY_TIME_COL = "Reply Time";
const REPLY_TEXT_COL = "Reply Text";
const UUID_COL = "UUID";
const EMAIL = Session.getEffectiveUser().getEmail();
 
/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mails')
      .addItem('Send Emails', 'sendEmails')
      .addItem('Reply', 'sendReplyToLastEmailForEachUUID')
      .addToUi();
}