/*
  
  Objective: This script used to post result updates on Facebook Page

*/

/* ------- Method to post on Facebook Page of Service ------- */
function postOnFB(FBpost) {
  var accessToken = "[FB_PAGE_ACCESS_TOKEN]";
  var url="https://graph.facebook.com/v2.2/me/feed";
  var response;
  
   var payload =
   {
     "message" : FBpost,
     "access_token" : accessToken
   };

   var options =
   {
     "method" : "post",
     "payload" : payload
   };
   
  while(true){
    try {
      response =  UrlFetchApp.fetch(url, options); 
      if(response.getResponseCode() === 200) {
        break;
      }
    } catch (err) {
      continue;
    }
  }
  Logger.log(response);
}

/* --------- Method to update old spreadheet content ----------*/
function setOldSpreadSheet(data2)
{
  var ss1 = SpreadsheetApp.openById('[Old_Spreadsheet_DocumentID]'); //old
  SpreadsheetApp.setActiveSpreadsheet(ss1);
  var sheet1 = SpreadsheetApp.getActiveSheet();
  
  for (var i = 0; i < data2.length; i++) 
  {
     for(var j = 0;j < data2[i].length;j++)
     { 
         sheet1.getRange(i+1,j+1).setValue(data2[i][j]);
     }
  }
}

/*--------- Updating results table in new Spreadhseet -----------*/

function insertIntoSpreadSheet(course,bname,sem) {
  
  var ss = SpreadsheetApp.openById('[New_Spreadsheet_DocumentID]'); //new
  SpreadsheetApp.setActiveSpreadsheet(ss);
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  var cell1 = 'A';
  var cell2 = 'B';
  var cell3 = 'C';
  
  for(var key in course)
  {
    var rowValue = (parseInt(key)+1);

    /*Initializing empty sem cells with value NA*/
    sheet.getRange('C'+rowValue).setValue("NA");
    sheet.getRange('D'+rowValue).setValue("NA");
    sheet.getRange('E'+rowValue).setValue("NA");
    sheet.getRange('F'+rowValue).setValue("NA");
    sheet.getRange('G'+rowValue).setValue("NA");
    sheet.getRange('H'+rowValue).setValue("NA");
    sheet.getRange('I'+rowValue).setValue("NA");
    sheet.getRange('J'+rowValue).setValue("NA");
    sheet.getRange('K'+rowValue).setValue("NA");
    sheet.getRange('L'+rowValue).setValue("NA");

    var cell1 = 'A'+rowValue;
    var cell2 = 'B'+rowValue;
    var cell3 = '';

    sheet.getRange(cell1).setValue(course[key]);
    sheet.getRange(cell2).setValue(bname[key]);
    
    /*creating a mapping from semester number to column number*/
    for(var key2 in sem[key])
    {

      switch(sem[key][key2].slice(0, sem[key][key2].indexOf(" ")))
      {
        case 'I': cell3='C';
                  break;
        case 'II': cell3='D';
                  break;
        case 'III': cell3='E';
                  break;
        case 'IV': cell3='F';
                  break;
        case 'V': cell3='G';
                  break;
        case 'VI': cell3='H';
                  break;
        case 'VII': cell3='I';
                  break;
        case 'VIII': cell3='J';
                  break;
        case 'IX': cell3='K';
                  break;
        case 'X': cell3='L';
                  break;
      }
      var cell3 = cell3+rowValue;

      sheet.getRange(cell3).setValue(sem[key][key2]);
    }
  }
}

/*------ Method to compare the Old & New Spreadhseet to check for new results -----*/
function compareSpreadSheets(){
  
  /* loading old spreadhseet */
  var ss1 = SpreadsheetApp.openById('[Old_Spreadsheet_DocumentID]'); //old
  SpreadsheetApp.setActiveSpreadsheet(ss1);
  var sheet1 = SpreadsheetApp.getActiveSheet();
  /* loading new spreadsheet */
  var ss2 = SpreadsheetApp.openById('[New_Spreadsheet_DocumentID]');  //new
  SpreadsheetApp.setActiveSpreadsheet(ss2);
  var sheet2 = SpreadsheetApp.getActiveSheet();

  /* getting data from both the spreasheet */
  var data1 = sheet1.getDataRange().getValues();
  var data2 = sheet2.getDataRange().getValues();
  
  var misMatch=false,change=[];
  var FBpost = "Results Declared !!";
  var updates=[];
  
  /* comparing each cell of spreadhseet for any updates */
  for (var i = 0; i < data2.length; i++) {
     if (typeof data1[i] == 'undefined') 
     {   
         misMatch = true;
         var post = data2[i][0]+" "+data2[i][1]+", Semester : ";
         for(var j = 2;j < data2[i].length;j++)
         { 
           if(data2[i][j]!='NA')
           {
              post=post + data2[i][j] + ",";
           }  
         }
         post=post.slice(0,post.length-1);
         
         updates.push(post);
     }
     else{
       var post2 = "";
       var pSem="";
       for(var j = 2;j < data2[i].length;j++)
       { 
         if(data1[i][j]!=data2[i][j]&&data2[i][j]!='NA')
         {
            pSem = pSem + data2[i][j] + ",";
            misMatch = true;
         }  
       }
       if(pSem.length>0)
       {
         post2 = post2 + data2[i][0] + " " + data2[i][1] + ", Semester : "+ pSem;
         post2=post2.slice(0,post2.length-1);
         
         updates.push(post2);
       }
     }
  }

  /*preparing string to post on Facebook Page*/
  for(var index in updates)
  {
    // putting \r\n instead of <br> to insert a line break
    FBpost = FBpost + "\r\n" + updates[index];
  }

  FBpost = FBpost + "\r\nTo get your result on email register on bit.do/manitrns\r\nor to view your result visit dolphintechnologies.in/manit\r\n#MRNS";
  Logger.log("Current Time : "+new Date().toLocaleTimeString());
  Logger.log("FBpost : "+FBpost);

  /* If there is a mismatch, then post the updates on Facebook Page*/
  if(misMatch)
  {
    Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Mistmatch Occcured ... Posting on FB");
    /*Posting on Facebook*/
    postOnFB(FBpost);
    Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Posted on FB ... setting Old Spreadsheet");
    /*Updating Old SpreadSheet*/
    setOldSpreadSheet(data2);
    Logger.log("Current Time : "+new Date().toLocaleTimeString()+" old SpreadSheet Setted .... Sending FB post email to the admin");
    /*Sending Notification email to admin regarding updates*/
    MailApp.sendEmail({
     to: "[Admin_Email_Address]",
     subject: "New FB Page Update",
     htmlBody: "I have posted the following on MRNS Page !!\r\n"+FBpost
    });
    Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Admin email successfully sent ... exiting compareSpreadsheet");
  }else
    Logger.log("No need to Post, Nothing new found !");
}

/*-------------------------------------*/

function knowDeclared() {
  Logger.log("Current Time : "+new Date().toLocaleTimeString());
  var time1 = new Date().getTime();
  while(true){
    try {
      var response =  UrlFetchApp.fetch("http://dolphintechnologies.in/manit/"); 
      if(response.getResponseCode() === 200) {
        break;
      }
    } catch (err) {
      continue;
    }
  }
  Logger.log("Current Time : "+new Date().toLocaleTimeString()+" HTTP Response received ... Now parsing it");
  var array=[];
  var html = response.getContentText();
  var string1 = "<marquee direction=\"up\" scrollamount=\"2\" height=\"205\" onMouseOver=\"this.stop();\" onMouseOut=\"this.start();\">";
  var string2 = "<marquee";
  var index1 = html.search(string2);
  html = html.slice(index1+string1.length);
  var string3 = "</marquee>";
  
  var index2 = html.search(string3);
  html = html.slice(0,index2);
  //till now the milk has been separated,now extracting the cream from it
  var string4 = "<p align=\"center\" class=\"style8\">";
  var string5 = "</p>";
  var l1;
  //html = html.replace(".","");
  //html = html.replace(" ","");
  while((l1=html.search(string4))>=0)
  {
    var l2 = html.search(string5) ;
    array.push(html.slice(l1+string4.length,l2+string5.length));
    html = html.replace(string4,"");
    html = html.replace(string5,":");
    html = html.replace("<strong>","<");
    html = html.replace("</strong>",">");
  }
  html = html.split("</p>");
  
  //clearing the stuff inside the array
  for(var element in array)
  {
    array[element] = array[element].replace(".","");
    array[element] = array[element].replace(" ","");
    array[element] = array[element].replace("<strong>","<");
    array[element] = array[element].replace("</strong>",">");
    array[element] = array[element].replace("</p>",":");
    array[element] = array[element].replace("&amp;","&");
    array[element] = array[element].replace(".","");
  }
  
  //now we have to just create a json array out of the array extracted above so as to store and easy access
  var text = '{"results":[ ';
  //Logger.log(array);
  //printing the array
  for(var element=0; element<array.length; element++)
  {
    var str = array[element];
    if((array[element])[0]==="<")
    { 
      var pos1 = 1;
      var pos2 = str.indexOf('>');
      var pos3 = str.indexOf('(');
      if(pos3>=0)
      {
        text+='{"course":"'+str.slice(pos1,pos2)+'","branch":[{"bname":"",';
        var pos4 = str.indexOf(')');
        str = str.slice(pos3+1);
        str = str.replace("(","");
        str = str.replace(")","");
        str = str.replace(":","");
        //str = str.replace("(","("");
        //str = str.replace(",",",");
        var res = str.split(",");
        for(var i in res)
        {
          res[i]='"'+res[i]+'"';
        }
        //Logger.log("Yes, It's Found :  "+res);
        text+= '"declared":['+res+']}]},';
      }
      else{
        text+='{"course":"'+str.slice(pos1,pos2)+'","branch":[';
        element+=1;
        while(element<array.length-1 && (array[element+1])[0]!="<")
        {
          var str2 = array[element];
          var p = str2.indexOf('(');
          text+='{"bname":"'+str2.slice(0,p)+'",';
          str2 = str2.slice(p+1);
          str2 = str2.replace("(","");
          str2 = str2.replace(")","");
          str2 = str2.replace(":","");
          var res = str2.split(",");
          for(var i in res)
          {
            res[i]='"'+res[i]+'"';
          }
          text+= '"declared":['+res+']},';
          element+=1;
        }
        text= text.slice(0,text.length-1);
        text+=']},';
      }
    }
  }
  text= text.slice(0,text.length-1);
  text+=']}';
  Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Response parsed successfully ... Converting into JSON and preparing FB post");
  var obj = JSON.parse(text);
  //Logger.log(out.results[i].course);
  //Logger.log(obj);
  var change = false;
  /*Reading Logic*/
  var doc = DocumentApp.openById('11xn7Fd9QyYqr1BFKhGBDcphyOyAPy1kclKUPRTHre20');
  var body = doc.getBody();
  var bodyText = body.getText();
  if(bodyText!=JSON.stringify(obj))
    change=true;
  /*variables for saving data in spreadsheet*/
  var decCourse=[],decBname=[],decSem=[];
  //Logger.log(decCourse[0]);
  //Logger.log(decBname);
  //Logger.log(decSem[0][2]);
  /*----------------------------------------*/
  for(var key1 in obj)
  {
    for(var key2 in obj[key1])
    {
      for(var key3 in obj[key1][key2]["branch"])
      {
        var output = obj[key1][key2]["course"]+" -> "+obj[key1][key2]["branch"][key3]["bname"]+" : "; 
        decCourse.push(obj[key1][key2]["course"]);
        decBname.push(obj[key1][key2]["branch"][key3]["bname"]);
        var tempSem=[];
        for(var key4 in obj[key1][key2]["branch"][key3]["declared"])
        {
          output+=obj[key1][key2]["branch"][key3]["declared"][key4]+", ";
          tempSem.push(obj[key1][key2]["branch"][key3]["declared"][key4]);
        }
        decSem.push(tempSem);
        //Logger.log(output);
      }
    }
  }
  Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Data converted into JSON successfully ... Inserting data into spreadsheet");
  insertIntoSpreadSheet(decCourse,decBname,decSem);
  Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Data inserted into the newDeclared spreadsheet successfully ... Comparing spreadsheets");
  compareSpreadSheets();
  Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Spreadsheets Compared successfully ... Saving JSON into the Word File")
  /*for(var key in decCourse)
  {
    Logger.log(decCourse[key]+" -> "+decBname[key]+" : "+decSem[key]);
  }*/
  if(change)
  {
    /*Writing Logic*/
    var text = body.editAsText();
    text.setText(''); // clear the current data from doc
    text.insertText(0, JSON.stringify(obj));  // insert at the beggining of the doc 
    Logger.log("There was a change");
  }else
    Logger.log("There was no change");
  Logger.log("Current Time : "+new Date().toLocaleTimeString()+" Script Finished successfully ...");  
  var time2 = new Date().getTime();
  if((time2-time1)/1000<60)
    Logger.log("Time Taken : "+(time2-time1)/1000+" seconds");
  else
    Logger.log("Time Taken : "+(time2-time1)/1000/60+" minutes");
}