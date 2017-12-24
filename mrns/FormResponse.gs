/*
  
  Objective: This script is used to send confirmation emails to the students after registration
             (It also send a follow up email containing the result, in case, result are available)

*/

/*---------Method to initialize the script trigger & email permissions-----------*/
function Initialize() {
 
  try {
 
    var triggers = ScriptApp.getProjectTriggers();
 
    for (var i in triggers)
      ScriptApp.deleteTrigger(triggers[i]);
 
    ScriptApp.newTrigger("EmailGoogleFormData").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onFormSubmit().create();
 
  } catch (error) {
    throw new Error("Please add this code in the Google Spreadsheet");
  }
}

/*Utility method to sleep for given number of seconds*/
function sleep(seconds) {
  var milliseconds = seconds*1000;
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds){
      break;
    }
  }
}

/* Utility Method to return spreadsheet list as JSON Object */
function getObjList() {
  var list = [];
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    obj["sch"] = String(data[i][1]);
    obj["sem"] = String(data[i][5]);
    obj["email"] = data[i][6];
    obj["status"] = data[i][8];
    list.push(obj);
  }
  return list; 
}

/* Method to update the status of a record as declared*/
function setDeclared(sch){
  var objList = getObjList();
  for (var i = 0; i < objList.length; i++) {
     if(objList[i].sch==sch){
        var cell = 'I'+(i+2);// +1, as index in spreadsheet starts from 1 not from 0 and +1, because we have taken looped data object from 1 not from zero
        SpreadsheetApp.getActiveSheet().getRange(cell).setValue('Y');
     }
  }
}

/* Method to update the status of a record as NOT declared */
function setUndeclared(sch){
  var objList = getObjList();
  for (var i = 0; i < objList.length; i++) {
     if(objList[i].sch==sch){
        var cell = 'I'+(i+2);// +1, as index in spreadsheet starts from 1 not from 0 and +1, because we have taken looped data object from 1 not from zero
        SpreadsheetApp.getActiveSheet().getRange(cell).setValue('N');
     }
  }
}

/* Method to update/set timestamp for a particular record identified by scholar number*/
function setTimeStamp(sch){
  var objList = getObjList();
  for (var i = 0; i < objList.length; i++) {
     if(objList[i].sch==sch){
        var cell = 'J'+(i+2); /* +1 as index in spreadsheet starts from 1 not from 0 and +1 because we have taken looped data object from 1 not from zero*/
        SpreadsheetApp.getActiveSheet().getRange(cell).setValue(new Date().toLocaleString());
     }
  }
}

/* Method to check for a result given scholar number & semester  */
function checkResult(sch,sem){
  
   var payload =
   {
     "scholar" : sch,
     "semester" : sem
   };

   var options =
   {
     "method" : "post",
     "payload" : payload
   };
  
   
  while(true){
    try {
      var response =  UrlFetchApp.fetch("http://dolphintechnologies.in/manit/accessview.php", options); 
      if(response.getResponseCode() === 200) {
        break;
      }
    } catch (err) {
      continue;
    }
  }
  
  var html = response.getContentText();
  
  html = html.replace("images/logoname.png","http://dolphintechnologies.in/manit/images/logoname.png");
  html = html.replace("images1/manit.png","http://dolphintechnologies.in/manit/images1/manit.png");
  
  /* regex to check for a result on website*/
  /* NOTE: I know that's too trivial, but that works fine */
  var status = html.match(/Scholar No. not found/g);

  /* setting latest time stamp to keep track on 'last checked' */
  setTimeStamp(sch);
    
  if(!status)
  {
    /* condition to handle a particular error case from the results page */
    var status2 = html.match(/SQL error/gi);

    if(status2)
    {
      return null;
    }
    /* If no error, then result is declared & is stored in html */
    return html;
  }  
  else
  {
    /* result no available */
    return null;
  }
  
}

/* Method to send email containing results*/
function sendResult(email,sch,html){
  
  var manitLogoUrl = 'http://dolphintechnologies.in/manit/images/logoname.png';
    MailApp.sendEmail({
      to: email,
      //bcc: "[Admin_Email_Address]", maybe used for moderation or testing or verfication
      subject: sch+" : Result Declared !!",
      htmlBody:html,
    });

  /* updating corrosponding record status as declared*/
  setDeclared(sch);
}

/* Method to handle Google Form Submission for Service */
function EmailGoogleFormData(e) {
  
  if (!e) {
    /*Please always run Initialize method before deploying script*/
    throw new Error("Please go the Run menu and choose Initialize");
  }
 
  try {
    
    if (MailApp.getRemainingDailyQuota() > 0) {
      
      /* If we have not totally exhausted daily email limit of Google Apps Script (free quota)*/
      var email = "",emailContent="";
      var sch,sem,name;
      var greetingTime = (new Date()).getHours();
      emailContent = getEmailContentFirst();

      var subject = "Form Submitted Successfully";
 
      var key, entry,
        message = "",
        ss = SpreadsheetApp.getActiveSheet(),
        cols = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
        
      /* Iterating through the Form Fields & preparing response email*/
      for (var keys in cols) {
 
        key = cols[keys];
        entry = e.namedValues[key] ? e.namedValues[key].toString() : "";
        if(key == "Email Address"){
          email+=entry;
        }
        else if(key == "Scholar Number"){
          sch=entry;
        }
        else if(key == "Semester"){
          sem=entry;
        }else if(key == "Name"){
          name=entry;
          name = name.substr(0,name.indexOf(' '));  //Getting only the first name
          name = name.charAt(0).toUpperCase()+name.slice(1).toLowerCase(); //making the first character uppercase while others lowercase
        }

        if(key!="Timestamp"){
          /* Only include form fields that are not blank */
          if ((entry !== "") && (entry.replace(/,/g, "") !== ""))
            message += getBeforeTd()+key+getBetweenTd()+entry+getAfterTd();
        }
      }

      emailContent = emailContent+"Hi "+name+",</h2>"; 
      emailContent = emailContent+"<p style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 16px;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;orphans: 3;widows: 3;font-weight: normal;margin-bottom: 20px;\"><span class=\"greeting\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;color: #36717e;\">"
      
      if(greetingTime<12){
        emailContent+= "Good Morning !! ";
      }else if(greetingTime<17){
        emailContent+= "Good Afternoon !! ";
      }else{
        emailContent+= "Good Evening !! ";
      }

      emailContent+="<\/span>You have Just Submitted the following information to us,<\/p><table class=\"details\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;border-spacing: 0;border-collapse: collapse;background-color: transparent;margin-left: 7%;width: 100% !important;\">";
      emailContent+=message;
      emailContent+="<\/table><p class=\"declaration\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 16px;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;orphans: 3;widows: 3;font-weight: normal;margin-bottom: 10px;margin-top: 10px;\">";
      
      var resultDeclared = false;
      var html = checkResult(sch,sem);
      
      if(html!==null){
        resultDeclared = true;
      }
       if(resultDeclared == false){
        emailContent+="Your Result is not yet Declared for the given semester.<br>But Don't worry We'll send it to you as soon as it is declared !!";
      }else{
        emailContent+="Your Result has already been Declared for the given semester.<br>You will get your result in the next email from us !!";
      }  

      emailContent+="</p>";
      emailContent+=getEmailContentSecond();
      
      /* just setting result undeclared for now (even if it's declared), will be handled in a moment*/
      setUndeclared(sch);

      /*------------- Logic to change color of Email Header & Contents ---------*/
      var hcolors = ["#0087CA","#4caf50","#009688","#cc8800","#3f51b5","#F44336"];
      var chosenColor = parseInt(sch)%6;
      var bgColorRegex =  new RegExp("#0087CA",'g');
      var colorRegex =  new RegExp("#36717e",'g');

      emailContent = emailContent.replace(bgColorRegex,hcolors[chosenColor]);
      emailContent = emailContent.replace(colorRegex,hcolors[chosenColor]);
      /*-------------logic end------------------*/

      
      MailApp.sendEmail({
      to: email,
      /* bcc: "[Admin_Email_Address]", may be used for moderation or verification */
      subject: subject,
      htmlBody:emailContent,
      });

      /* sending the result after waiting for 22 seconds*/
      /* yeah! it's just 22 no logic behind it (actually my scholar/registration number) :p */
      sleep(22);
      if(resultDeclared){
        sendResult(email,sch,html);
      }
    }
  } catch (error) {
    Logger.log(error.toString());
  }
}

/* Methods to return HTML content of reponse email as strings */
/* Obviously, it's not the best way to handle this, maybe some files/spreadsheets could be used,
   or even better some templating engine may be used */

function getEmailContentFirst(){
  var content="<!DOCTYPE html>\r\n<html xmlns=\"http:\/\/www.w3.org\/1999\/xhtml\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 10px;font-family: sans-serif;line-height: 1.65;-webkit-text-size-adjust: 100%;-ms-text-size-adjust: 100%;-webkit-tap-highlight-color: rgba(0,0,0,0);\">\r\n<head style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;\">\r\n    <meta http-equiv=\"Content-Type\" content=\"text\/html; charset=utf-8\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;\">\r\n    <meta name=\"viewport\" content=\"width=device-width\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;\">\r\n\r\n    <link rel=\"stylesheet\" href=\"http:\/\/maxcdn.bootstrapcdn.com\/bootstrap\/3.3.6\/css\/bootstrap.min.css\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;\">\r\n\r\n    <!-- jQuery library -->\r\n    <script src=\"https:\/\/ajax.googleapis.com\/ajax\/libs\/jquery\/1.12.0\/jquery.min.js\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;\"><\/script>\r\n\r\n    <!-- Latest compiled JavaScript -->\r\n    <script src=\"http:\/\/maxcdn.bootstrapcdn.com\/bootstrap\/3.3.6\/js\/bootstrap.min.js\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;\"><\/script>\r\n    <!--link href=\'https:\/\/fonts.googleapis.com\/css?family=Open+Sans\' rel=\'stylesheet\' type=\'text\/css\'-->\r\n    \r\n<\/head>\r\n<body style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 14px;font-family: &quot;Helvetica Neue&quot;,Helvetica,Arial,sans-serif;line-height: 1.42857143;color: #333;background-color: #fff;height: 100%;background: #efefef;-webkit-font-smoothing: antialiased;-webkit-text-size-adjust: none;width: 100% !important;\">\r\n<table class=\"body-wrap\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;border-spacing: 0;border-collapse: collapse;background-color: transparent;height: 100%;background: #efefef;-webkit-font-smoothing: antialiased;-webkit-text-size-adjust: none;width: 100% !important;\">\r\n    <tr style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;page-break-inside: avoid;\">\r\n        <td class=\"container\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0 auto !important;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;padding-right: 15px;padding-left: 15px;margin-right: auto;margin-left: auto;display: block !important;clear: both !important;max-width: 580px !important;\">\r\n\r\n            <!-- Message start -->\r\n            <table style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;border-spacing: 0;border-collapse: collapse;background-color: transparent;width: 100% !important;\">\r\n                <tr style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;page-break-inside: avoid;\">\r\n                    <td align=\"center\" class=\"masthead\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 80px 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;background: #0087CA;color: white;\">\r\n\r\n                        <h1 style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0 auto !important;padding: 0;font-size: 32px;font-family: inherit;line-height: 1.25;font-weight: 500;color: inherit;margin-top: 20px;margin-bottom: 20px;max-width: 90%;text-transform: uppercase;\">M.A.N.I.T. Results Notification Service<\/h1>\r\n\r\n                    <\/td>\r\n                <\/tr>\r\n                <tr style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;page-break-inside: avoid;\">\r\n                    <td class=\"content\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 30px 35px;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;background: white;\">\r\n\r\n                        <h2 style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 28px;font-family: inherit;line-height: 1.25;orphans: 3;widows: 3;page-break-after: avoid;font-weight: 500;color: inherit;margin-top: 20px;margin-bottom: 20px;\">";
  return content;
}

function getEmailContentSecond(){
  var content="<\/td><\/tr>\r\n<\/table><\/td>\r\n<\/tr>\r\n    <tr style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;page-break-inside: avoid;\">\r\n        <td class=\"container\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0 auto !important;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;padding-right: 15px;padding-left: 15px;margin-right: auto;margin-left: auto;display: block !important;clear: both !important;max-width: 580px !important;\">\r\n\r\n            <!-- Message start -->\r\n            <table style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;border-spacing: 0;border-collapse: collapse;background-color: transparent;width: 100% !important;\">\r\n                <tr style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;page-break-inside: avoid;\">\r\n                    <td class=\"content footer\" align=\"center\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 30px 35px;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;background: none;\">\r\n                        <p style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 14px;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;orphans: 3;widows: 3;font-weight: normal;margin-bottom: 0;color: #888;text-align: center;\">Thanks for using M.A.N.I.T. Results Notification Service. \r\n                            For any assistance or queries drop an email to <a href=\"mailto:manitresults@gmail.com\" style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;background-color: transparent;color: #888;text-decoration: none;font-weight: bold;\">manitresults@gmail.com<\/a><\/p>\r\n                    <\/td>\r\n                <\/tr>\r\n            <\/table>\r\n\r\n        <\/td>\r\n    <\/tr>\r\n<\/table>\r\n<\/body>\r\n<\/html>";
  return content;
}

function getBeforeTd(){
  content="<tr style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 100%;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;page-break-inside: avoid;\"><td style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 1.15em;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;height: 28px;vertical-align: bottom;\">";
  return content;
}

function getBetweenTd(){
  content="<\/td><td style=\"-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;margin: 0;padding: 0;font-size: 1.15em;font-family: \'Avenir Next\', &quot;Helvetica Neue&quot;, &quot;Helvetica&quot;, Helvetica, Arial, sans-serif;line-height: 1.65;height: 28px;vertical-align: bottom;\">";
  return content;
}

function getAfterTd(){
  content="</td></tr>";
  return content;
}