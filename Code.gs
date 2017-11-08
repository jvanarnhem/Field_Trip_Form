// Replace with your spreadsheet ID
var ss = SpreadsheetApp.openById("1I69c_RTENUx-VNfkeX6e39GHYpQr-g1yU0367OFWeRg");
var SheetRecApps = ss.getSheetByName("RecentApplications");

//***********************************************************************************************************************************************

function doGet(e) {
  var buildingApproved = e.parameter.buildapprove;
  var appNumber = e.parameter.appnum;
  
  if (buildingApproved == 2) {
    var template = HtmlService.createTemplateFromFile("DistrictAdmin.html");
    var data = SheetRecApps.getDataRange().getValues();
    
    for (var i in data) {
      if (data[i][0] == appNumber) {
        template.app = data[i];
        break;
      }
    };
    //template.comments = form.buildcomments;
  } else if (buildingApproved == 1){
    var template = HtmlService.createTemplateFromFile("BuildAdmin.html");
    var data = SheetRecApps.getDataRange().getValues();
    
    for (var i in data) {
      if (data[i][0] == appNumber) {
        template.app = data[i];
        break;
      }
    };
  } else if (buildingApproved == 0) {
    update(appNumber, 0);
    return ContentService.createTextOutput("Rejected!");
  } else {
    var template = HtmlService.createTemplateFromFile("Form.html");
  }
  
  var html = template.evaluate();
  return HtmlService.createHtmlOutput(html);
}
//***********************************************************************************************************************************************

function isValidEmail_(email) {
  var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(email);
}
//***********************************************************************************************************************************************

function postApplication(form) {
  // validate user email
  if (!isValidEmail_(form.email))
    throw "Please provide a valid email address.";
  if (form.destaddress == null || form.destaddress == "")
    throw "Please enter a full address for your destination.";
  if (form.leaveschool == null || form.leaveschool == "")
    throw "Please enter a time for your school departure.";
  if (form.arrivedestination == null || form.arrivedestination == "")
    throw "Please enter a time for arrival at your destination.";
  if (form.leavedestination == null || form.leavedestination == "")
    throw "Please enter a time for departure from your destination.";
  if (form.arriveschool == null || form.arriveschool == "")
    throw "Please enter a time for arrival at school.";
  
  var appNumber = +new Date();
  var timeStamp = new Date();
  timeStamp = "'" + timeStamp.toLocaleDateString('en-US') ;
  var myDateArray = form.tripdate.split("-");
  var dateOfTrip = new Date(myDateArray[0],myDateArray[1]-1,myDateArray[2]);
  var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var dayOfWeek = days[dateOfTrip.getDay()];
  dateOfTrip = "'" + dateOfTrip.toLocaleDateString('en-US');
  var leave1 = form.leaveschool.toHHMM();
  var arrive1 = form.arrivedestination.toHHMM();
  var leave2 = form.leavedestination.toHHMM();
  var arrive2 = form.arriveschool.toHHMM();

  var buildadminemail = "";
  var buildadminname = "";
  if (form.building == "HS") {
    buildadminemail = "hschafer@ofcs.net";
    buildadminname = "Holly Schafer";
  } else if (form.building == "MS") {
    buildadminemail = "mkurz@ofcs.net";
    buildadminname = "Mark Kurz";
  } else if (form.building == "IS") {
    buildadminemail = "dsvec@ofcs.net";
    buildadminname = "Don Svec";
  } else if (form.building == "FL") {
    buildadminemail = "lbarrett@ofcs.net";
    buildadminname = "Lisa Barrett";
  } else {
    buildadminemail = "mbrunner@ofcs.net";
    buildadminname = "Melinda Brunner";
  }
  
  // Construct form element value in an array
  var application = [
    appNumber,
    timeStamp,
    form.destination,
    dateOfTrip,
    dayOfWeek,
    form.building,
    form.adultincharge,
    form.email,
    form.phone,
    form.adultassisting,
    form.numadults,
    form.numstudents,
    form.largebus,
    form.smallbus,
    form.vans,
    form.depart,
    form.destaddress,
    leave1,
    arrive1,
    leave2,
    arrive2,
    form.eat,
    form.restroom,
    form.comments,
    "",
    buildadminname
  ];
  
  SheetRecApps.appendRow(application);
  
  update(appNumber, 1);
  var htmlBody = "<p>The following application was submitted for approval.</p>";
  htmlBody += "<p>Destination: " + form.destination + "</p>";
  htmlBody += "<p>Date: " + dateOfTrip + "</p>";
  htmlBody += "<p>Adult/Teacher in Charge: " + form.adultincharge + "</p>";
  
  MailApp.sendEmail({
    to: form.email,
    subject: "Field trip application confirmation for " + form.tripdate,
    htmlBody: htmlBody + "<p>Please reference the application number: " + appNumber + " in any correspondence with the bus garage or administrator.</p>"
  });
  
  htmlBody += "<p>&nbsp;</p>";
  htmlBody += '<p>Click <a href="' + ScriptApp.getService().getUrl()
     + '?appnum=' + appNumber + '&buildapprove=1' 
     + '">here</a> to see details and approve/reject.</p>';
  
  Logger.log(buildadminemail);
  // CHANGE THIS TO BUILDING ADMINISTRATOR
  MailApp.sendEmail({
    to: "javanarnhem@gmail.com",
    subject: "Action Required: Field Trip Application from " + form.adultincharge,
    htmlBody: htmlBody
  });
  
  // Return confirmation message to user
  return "Application placed successfully. \nYou should receive a confirmation email, and more details have been sent to your building administrator for approval. \nYou may close this window now.";
}
//***********************************************************************************************************************************************

function finalApproval(form) {
  var timeStamp = new Date();
  var appNumber = form.appnum;
  
  update(appNumber, 2);
  
  var data = SheetRecApps.getDataRange().getValues();
    
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == appNumber) {
      SheetRecApps.getRange(i+1, 27, 1, 1).setValue(form.buildcomments);
      break;
    }
  };
  
  var htmlBody = "<p>The following application was submitted for approval.</p>";
  htmlBody += "<p>Destination: " + form.destination + "</p>";
  htmlBody += "<p>Date: " + form.tripdate + "</p>";
  htmlBody += "<p>Adult/Teacher in Charge: " + form.adultincharge + "</p><br />";
  htmlBody += "<p>Building Administrator Comments: " + form.buildcomments + "</p>";
  htmlBody += '<p>Click <a href="' + ScriptApp.getService().getUrl()
     + '?appnum=' + appNumber + '&buildapprove=2' 
     + '">here</a> to see details and approve/reject.</p>';
  
  // CHANGE TO KCOGAN@OFCS.NET
  MailApp.sendEmail({
    to: "jvanarnhem@ofcs.net",
    subject: "Action Required: Field Trip Application from " + form.adultincharge,
    htmlBody: htmlBody,
  });
  
  return "The field trip application will now be submitted for district administrator approval. You may close this window now.";
}
//***********************************************************************************************************************************************

function approvalComplete (form) {
  var timeStamp = new Date();
  var appNumber = form.appnum;
  
  update(appNumber, 3);
  
  var data = SheetRecApps.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == appNumber) {
      SheetRecApps.getRange(i+1, 28, 1, 1).setValue(form.districtcomments);
      break;
    }
  };
  
  var htmlBody = "<p>The following application has been approved.</p>";
  htmlBody += "<p>Destination: " + form.destination + "</p>";
  htmlBody += "<p>Date: " + form.tripdate + "</p><br />";
  htmlBody += "<p>Building Administrator Comments: " + form.buildcomments + "</p>";
  htmlBody += "<p>District Administrator Comments: " + form.districtcomments + "</p>";
  
  var finishedDoc = doMerge(appNumber, form.adultincharge, form.destination, form.building);
  addToCalendar(form);
  MailApp.sendEmail({
    to: form.email,
    subject: "Field Trip Approval Notification",
    htmlBody: htmlBody + "<p>Please reference the application number: " + appNumber + " in any correspondence with the bus garage or administrator.</p>"
  });
  
  moveApproved();
  
  var recipients = "javanarnhem@gmail.com, hkrakowiak@ofcs.net, falls1@ofcs.net";
  var htmlBody2 = "Attached is the field trip application document submitted by " + form.adultincharge;
  MailApp.sendEmail(recipients,"Field Trip Approval Notification", htmlBody2,
  {
      attachments: [finishedDoc.getAs(MimeType.PDF)]
  });
  
  return "Approval is complete. Details will be sent to the submitter and the bus garage. You may close this window now.";
}
//***********************************************************************************************************************************************

function rejection1 (form) {
  var appNumber = form.appnum;
  update(appNumber, 0);
  
  var htmlBody = "<p>Sorry, your field trip application has been rejected by <br />";
  htmlBody += "your building administrator.</p>";
  htmlBody += "<p>Building administrator comments: " + form.buildcomments + "</p>";
  htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
  htmlBody += '<p>Click <a href="' + ScriptApp.getService().getUrl() 
     + '">here</a> to resubmit your application.</p>';
     
  MailApp.sendEmail({
    to: form.email,
    subject: "Please read: Field trip application information",
    htmlBody: htmlBody
  });
  
  return "A notice of this rejection will be sent to the applicant. You may close this window now.";
}

function rejection2 (form) {
  var appNumber = form.appnum;
  update(appNumber, 0);
  
  var htmlBody = "<p>Sorry, your field trip application has been rejected by <br />";
  htmlBody += "the district administrator.</p>";
  htmlBody += "<p>District administrator comments: " + form.districtcomments + "</p>";
  htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
  htmlBody += '<p>Click <a href="' + ScriptApp.getService().getUrl() 
     + '">here</a> to resubmit your application.</p>';
     
  MailApp.sendEmail({
    to: form.email,
    subject: "Please read: Field trip application information",
    htmlBody: htmlBody
  });
  
  return "A notice of this rejection will be sent to the applicant. You may close this window now.";
}
//***********************************************************************************************************************************************

function update (num, val) {
  var data = SheetRecApps.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
      if (data[i][0] == num) {
        if (val == 3) {
          SheetRecApps.getRange(i+1, 25, 1, 1).setValue("APPROVED").setBackground("#00FF00");
        } else if (val == 2) {
          SheetRecApps.getRange(i+1, 25, 1, 1).setValue("Pending Final Approval").setBackground("#FFFF00");
        } else if (val == 1) {
          SheetRecApps.getRange(i+1, 25, 1, 1).setValue("Pending Building Approval").setBackground("#00FFFF");
        } else {
          SheetRecApps.getRange(i+1, 25, 1, 1).setValue("Rejected").setBackground("red");
        }
        break;
      }
  };
}
//***********************************************************************************************************************************************

function moveApproved() {
  // moves a row from a sheet to another when a magic value is entered in a column
  // adjust the following variables to fit your needs
  // see https://productforums.google.com/d/topic/docs/ehoCZjFPBao/discussion

  var sheetNameToWatch = "RecentApplications";

  var columnNumberToWatch = 25; // column A = 1, B = 2, etc.
  var valueToWatch = "APPROVED";
  var sheetNameToMoveTheRowTo = "Completed";
  
  var data = SheetRecApps.getDataRange().getValues();
  //data.shift();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][24] == valueToWatch) {
      var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
      var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
      SheetRecApps.getRange(i+1, 1, 1, SheetRecApps.getLastColumn()).moveTo(targetRange);
      SheetRecApps.deleteRow(i+1);
    }
  };
}
//***********************************************************************************************************************************************

String.prototype.toHHMM = function () {
    var tag = "";
    var timeArray = this.split(":");
    var hour = parseInt(timeArray[0],10);
    var minute = parseInt(timeArray[1],10);
    if (hour >=12) {
      tag = " pm";
      if (hour > 12) {hour = hour - 12;}
    }
    else {
      tag = " am";
    }
    if (minute < 10) {
       minute = "0" + minute;
    }

    return "'" + hour + ':' + minute + tag;
}