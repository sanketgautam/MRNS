/*
  
  Objective: This script regularly checks for result updates for each record after certain time,
             and send update email to the corrosponding students

*/

/* Utility Method to return spreadsheet list as JSON Object */
function getObjList() {
  var list = []; 
  var ss = SpreadsheetApp.openById('[Registration_Form_Spreadsheet_DocumentID]');

  SpreadsheetApp.setActiveSpreadsheet(ss);
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
      var obj = {};
      obj["sch"] = String(data[i][1]);
      obj["sem"] = String(data[i][5]);
      obj["email"] = data[i][6];
      obj["status"] = data[i][8];
      obj["lastChecked"] = data[i][9];
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

/* Method to send email containing results*/
function sendResult(email,sch,html){
  var manitLogoUrl = 'http://dolphintechnologies.in/manit/images/logoname.png';
    MailApp.sendEmail({
      to: email,
      //bcc: "[Admin_Email_Address]",
      subject: sch+" : Result Declared !!",
      htmlBody:html,
    });
  setDeclared(sch);
}

/* Method to update/set timestamp for a particular record identified by scholar number*/
function setTimeStamp(sch){
  var objList = getObjList();
  for (var i = 0; i < objList.length; i++) {
     if(objList[i].sch==sch){
        var cell = 'J'+(i+2);// +1, as index in spreadsheet starts from 1 not from 0 and +1, because we have taken looped data object from 1 not from zero
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

/* Method to regularly check for result updates for stored records */
function checkAndSendResults() {
  var time1 = new Date().getTime();
  Logger.log("Current time is : "+new Date().toLocaleTimeString());
  
  var counter=0;  
  var objList = getObjList();

  Logger.log("Total Number of records : "+String(objList.length));
  
  /*Logic to set the value of i such that it will scan only half of the records at a time*/
  /* Just a very trivial way to handle the HTTPRequest limit of Google Apps Script (free quota) 
     some better approach should be used*/

  var cell = 'J1';
  var val = SpreadsheetApp.getActiveSheet().getRange(cell).getValue();
  var i;
  if(val=="even")
  {  
    SpreadsheetApp.getActiveSheet().getRange(cell).setValue("odd");
    i=0;
  }
  else{
    SpreadsheetApp.getActiveSheet().getRange(cell).setValue("even");
    i=1;
  }
  /*-----------------value of i set successfully,now it's time to use it----------------*/

  for (; i < objList.length; i+=2) {
     if(objList[i].status=='N'){
       var html = checkResult(objList[i].sch,objList[i].sem);
       if(html!==null){
         sendResult(objList[i].email,objList[i].sch,html);
       }
       var sch = parseInt(objList[i].sch);

       Logger.log(String(sch));
       var date = new Date(objList[i].lastChecked);

       counter++;
     }
  }
  /* Logging total number of records processed & method latency */
  Logger.log("Number of records processed : "+counter);
  var time2 = new Date().getTime();
  if((time2-time1)/1000<60)
    Logger.log("Time Taken : "+(time2-time1)/1000+" seconds");
  else
    Logger.log("Time Taken : "+(time2-time1)/1000/60+" minutes");
}