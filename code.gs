function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('1. Fetch the Email Subject', 'initialize')
    .addItem('2. Set Schedule for Emails', 'setSchedule')
    // .addItem('Send Emails', 'sendMails')
    .addItem('Check Your Email Quota', 'checkQuota')
    .addToUi();
}
function initialize() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(2, 1, sheet.getLastRow() + 1, 10);
  /* Delete all existing triggers */
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendMails") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  /* Import Gmail Draft Messages into the Spreadsheet */
  var drafts = GmailApp.getDraftMessages();
  // Logger.log(drafts.length);
  // return;
  // Logger.log(drafts);
  if (drafts.length > 0) {
    var rows = [];
    for (var i = 0; i < drafts.length; i++) {
      if (drafts[i]) {
        rows.push([drafts[i].getId(), drafts[i].getSubject()]);
        Logger.log(drafts[i].getId());
      }
    }
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}



/* Create time-driven triggers based on Gmail send schedule */
function setSchedule() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var time = new Date().getTime();
  ScriptApp.newTrigger("sendMails")
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log(data);
  var code = [];
  for (var row in data) {
    if (row != 0) {
      var schedule = data[row][3];
      // Logger.log(schedule.getTime());
      // return;

      if (schedule !== "") {
        if (schedule.getTime() > time) {
          // ScriptApp.newTrigger("sendMails")
          //   .timeBased()
          //   .at(schedule)
          //   .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
          //   .create();
          code.push("Scheduled");

        } else {
          code.push("Date is in the past");
        }
      } else {
        code.push("Not Scheduled");
      }
      // Logger.log(code);
    }
  }
  for (var i = 0; i < code.length; i++) {
    sheet.getRange("E" + (i + 2)).setValue(code[i]);
  }
}

function checkSchedule(scheduledTime) {
  let theCurrentTime = new Date().getTime();
  if (scheduledTime <= theCurrentTime) {
    return true;
  }
}

const RECIPIENT_COL = "Recipients";
const EMAIL_SENT_COL = "Email Sent";
const EMAIL_STATUS_COL = "Email Status";

function sendMails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {

  var data = sheet.getDataRange().getDisplayValues();

  const heads = data.shift();

  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const recipientColIdx = heads.indexOf(RECIPIENT_COL);
  const emailStatusColIdx = heads.indexOf(EMAIL_STATUS_COL);

  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const out = [];
  const messageRes = [];

  let time = new Date().getTime();

  let emailStat = [];

  obj.forEach(function (row, rowIdx) {

    subjectLine = data[rowIdx][1];
    const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
    // if (data[rowIdx][4] == "Delivered") {
    //   return;
    // } else if (data[rowIdx][4] != "Scheduled") {
    //   Logger.log("Not Scheduled");
    //   return;
    // } else
    let nameOfTheSender = data[rowIdx][6];
      let schedule = data[rowIdx][3];
      let scheduledTime = Date.parse(schedule);
      let emailAddress = data[rowIdx][2];
     if (data[rowIdx][4] == "Scheduled" && scheduledTime <= time) {
      
      // if (scheduledTime > time) {
      //   Logger.log("The times didn't match");
      //   // emailStat.push("Delayed");
      //   return;
      // }
      const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
      GmailApp.sendEmail(emailAddress, msgObj.subject, msgObj.text, {
             htmlBody: msgObj.html,
             // bcc: 'a.bbc@email.com',
             // cc: 'a.cc@email.com',
             // from: 'an.alias@email.com',
             name: nameOfTheSender,
             // replyTo: 'a.reply@email.com',
             // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
             attachments: emailTemplate.attachments,
             inlineImages: emailTemplate.inlineImages
           });
      Logger.log("To: " + emailAddress);
      Logger.log("Subj: " + msgObj.subject);
      Logger.log("Body: " + msgObj.text);
      Logger.log("Sender: " + nameOfTheSender);
      emailStat.push("Delivered");
      sheet.getRange("E" + (rowIdx+2)).setValue(emailStat);


    } else {
      Logger.log("No Conditions Matched")
      return;
    }




  });

  // for (var i = 0; i < emailStat.length; i++) {
  //   sheet.getRange("E" + (i + 2)).setValue(emailStat[i]);
  // }

}

function getGmailTemplateFromDrafts_(subject_line) {
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    // get the message object
    const msg = draft.getMessage();

    // Handling inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Get all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
    const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
    const htmlBody = msg.getBody();

    // Create an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

    //Regexp to search for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiate the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {
      message: { subject: subject_line, text: msg.getPlainBody(), html: htmlBody },
      attachments: attachments, inlineImages: inlineImagesObj
    };
  } catch (e) {
    throw new Error("Oops - can't find Gmail draft");
  }

  /**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} GmailDraft object
  */
  function subjectFilter_(subject_line) {
    return function (element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

function fillInTemplateFromObject_(template, data) {
  // we have two templates one for plain text and the html body
  // stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}


function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};

function checkQuota() {
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getUi().alert("Remaining email quota (refreshes daily): " + emailQuotaRemaining);
}
