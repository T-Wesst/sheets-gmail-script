var EMAIL_SENT = 'EMAIL_SENT';

function include(filename) {
return HtmlService.createHtmlOutput(filename).getContent();
}

function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  const numOfRecipients = SpreadsheetApp.getActiveSheet().getRange('H2').getValue();
  var startCol = 1;
  var endCol = 7; 
  
  
  var htmlTemplate = HtmlService.createTemplateFromFile('index');
     
  
  var dataRange = sheet.getRange(startRow, startCol, numOfRecipients, endCol);
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var recipientName = row[1];
    var emailAddress = row[3];
    var emailSent = row[7]; 

    var subject = "Your Email Subject Here";
    htmlTemplate.recipientName = recipientName;
    
  var htmlBody = htmlTemplate.evaluate().getContent();
    var user = Session.getActiveUser().getEmail();
    if (emailSent != EMAIL_SENT) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: htmlBody
      });
      SpreadsheetApp.flush();
    }
  }
}
