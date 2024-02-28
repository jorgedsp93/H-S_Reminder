function sendTodayTalk() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workers Responses');
  var lastRow = sheet.getLastRow();
  var emails = sheet.getRange(5, 2, lastRow - 4, 1).getValues();
  var names = sheet.getRange(5, 1, lastRow - 4, 1).getValues();
  var dates = sheet.getRange(1, 3, 1, lastRow - 2).getValues()[0];
  var topics = sheet.getRange(2, 3, 1, lastRow - 2).getValues()[0];
  var links = sheet.getRange(4, 3, 1, lastRow - 2).getValues()[0];
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (var col = 0; col < dates.length; col++) {
    var date = new Date(dates[col]);
    date.setHours(0, 0, 0, 0);
    if (date.getTime() != today.getTime()) continue;
    
    var topic = topics[col];
    var link = links[col];
    
    for (var row = 0; row < emails.length; row++) {
      var email = emails[row][0];
      if (!email || email.trim() === '') continue; // Skip if no email address
      
      var name = names[row][0];
      
      var subjectLine = 'Health and Safety Talk - ' + topic;
      var emailBody = 'Good Morning ' + name + '!\n\n' +
                      'For today\'s Health & Safety talk, in order to comply with an MOL requirement we will review ' + topic + '.\n\n' +
                      'Remember to write your name and check the acknowledgment box at the end of the form,\n\n' +
                      link + '\n\n' +
                      'Regards,\n' +
                      'Jorge Serrano,';
      
      MailApp.sendEmail(email, subjectLine, emailBody);
    }
  }
}
