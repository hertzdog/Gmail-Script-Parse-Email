/*

Based on: https://gist.github.com/Ferrari/9678772
and on: http://mikecr.it/ramblings/marking-gmail-read-with-apps-script
and on: https://ctrlq.org/code/20019-parse-gmail-extract-data

*/


function parseEmailMessages() {

  //start = start || 0;

  //var threads = GmailApp.getInboxThreads(start, 100);
  //var threads = GmailApp.getUserLabelByName("varie-iscrizioni").getThreads()
  var ss = SpreadsheetApp.create("test");
  var sheet = ss.getActiveSheet();
  var threads = GmailApp.search('label:varie-iscrizioni label:unread');

  sheet.appendRow(["date", "time","day", "name", "surname", "email", "mobile", "nation", "postalcode", "", "", "page from"]);

  for (var i = 0; i < threads.length; i++) {

    // Get the first email message of a threads
    var tmp,
      message = threads[i].getMessages()[0],
      subject = message.getSubject(),
      content = message.getPlainBody(),
        d = message.getDate();

    // Get the plain text body of the email message
    // You may also use getRawContent() for parsing HTML

    // Implement Parsing rules using regular expressions
    if (content) {

      tmp = content.match(/Nome;\s*([A-Za-z0-9\s]+)(\r?\n)/);
      var name = (tmp && tmp[1]) ? tmp[1].trim() : '';

      tmp = content.match(/Cognome;\s*([A-Za-z0-9\s]+)(\r?\n)/);
      var surname = (tmp && tmp[1]) ? tmp[1].trim() : '';

      tmp = content.match(/email;\s*([A-Za-z0-9@.]+)/);
      var email = (tmp && tmp[1]) ? tmp[1].trim() : '';

      tmp = content.match(/cellulare;\s*([\+A-Za-z0-9@.\-]+)/);
      var mobile = (tmp && tmp[1]) ? tmp[1].trim() : '';

      tmp = content.match(/nazione;\s*([A-Za-z0-9\s]+)(\r?\n)/);
      var nation = (tmp && tmp[1]) ? tmp[1].trim() : '';

      tmp = content.match(/cap;\s*([A-Za-z0-9\s]+)(\r?\n)/);
      var postalcode = (tmp && tmp[1]) ? tmp[1].trim() : '';

      tmp = content.match(/Pagina;\s*([\w\-]+)(\r?\n)/);
      var page = (tmp && tmp[1]) ? tmp[1].trim() : '';

      var day = d.getDate()  + "/" + (d.getMonth()+1) + "/" + d.getFullYear();
      var time = ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);


      sheet.appendRow([d, time, day, name, surname, email, mobile, nation, postalcode, "", "", page]);

      // set the thread read
      GmailApp.markThreadRead(threads[i]);

    } // End if

  } // End for loop
}
