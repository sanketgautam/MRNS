/*

  Objective: This script is used to send email about new notices/circulars on the website

  Note:  This script was designed to work on notices page of old website,
         to make it work on new wesbite, some modifications may be required.

*/

function myNoticeUpdates() {
  
  /* fetching content from the notices page of the wesbite */ 
  var response = UrlFetchApp.fetch("http://www.manit.ac.in/manitbpl/index.php?option=com_content&view=article&id=92&Itemid=182");
  var requestPage = response.getContentText();
  var str1 = "<span style=\"color: #0000ff;\">Other Notices</span>";
  var str2 = "<span style=\"color: #0000ff;\">PhD Seminar</span>";
  
  /* extracting all content between str1 and str2 in notices */
  var location1  = requestPage.search(str1);
  var length1 = (str1).length;
  var length2 = (str2).length;
  
  var notices = requestPage.slice(location1+length1);
  var location2 = notices.search(str2);
  notices = notices.slice(0,location2+length2);
  var length = notices.length;
  
  /* converting all relative links to absolute links on the page */
  notices = notices.replace(/Year/g,"http://www.manit.ac.in/manitbpl/Year");
  notices = notices.replace(/year/g,"http://www.manit.ac.in/manitbpl/Year");
  notices = notices.replace(/index.php/g,"http://www.manit.ac.in/manitbpl/index.php");
  
  /* opening document containing previous image/content of page */
  var doc = DocumentApp.openById('Google_Drive_DocumentID');
  var body = doc.getBody();

  /* extracting text from the document */
  var bodyText = body.getText();

  /* Checking whether there is any change among the two states of page */
  if(bodyText == notices){
    Logger.log("No new Updates");
  }
  else
  {
    Logger.log("New Updates Available");
    var text = body.editAsText();
    text.setText(''); /* clear the current data from doc */
    text.insertText(0, notices);  /* insert new content at the beggining of the doc */
    
    /* sending notification email about changes in the website */
    MailApp.sendEmail({
     to: "[Admin_Email_Address]",
     subject: "New Update From MANIT Website",
     htmlBody: notices
    });
  }
  Logger.log(notices);
}