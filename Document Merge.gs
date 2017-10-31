/*  Code source: https://github.com/hadaf
 *  This is the main method that should be invoked. 
 *  Copy and paste the ID of your template Doc in the first line of this method.
 *
 *  Make sure the first row of the data Sheet is column headers.
 *
 *  Reference the column headers in the template by enclosing the header in square brackets.
 *  Example: "This is [header1] that corresponds to a value of [header2]."
 */
function doMerge(appNumber, adultcharge, end, building) {
  var selectedTemplateId = "1umOfWfXGnHU8qShPADRHGA-1LPg9bHmz6xhxnR1dDx8";//Copy and paste the ID of the template document here (you can find this in the document's URL)
  
  var templateFile = DriveApp.getFileById(selectedTemplateId);
  var targetFolder = DriveApp.getFolderById("0BwPBfS-vilmdVGJwWWZybzlsdjQ");
  var mergedFile = templateFile.makeCopy(targetFolder); //make a copy of the template file to use for the merged File. 
  // Note: It is necessary to make a copy upfront, and do the rest of the content manipulation inside this single copied file, 
  // otherwise, if the destination file and the template file are separate, a Google bug will prevent copying of images from the 
  // template to the destination. See the description of the bug here: https://code.google.com/p/google-apps-script-issues/issues/detail?id=1612#c14
  mergedFile.setName(appNumber + " Field Trip Application for " + adultcharge);//give a custom name to the new file (otherwise it is called "copy of ...")
  var mergedDoc = DocumentApp.openById(mergedFile.getId());
  var bodyElement = mergedDoc.getBody();//the body of the merged document, which is at this point the same as the template doc.
  var bodyCopy = bodyElement.copy();//make a copy of the body
  
  bodyElement.clear();//clear the body of the mergedDoc so that we can write the new data in it.
  
  var sheet = SpreadsheetApp.openById("1I69c_RTENUx-VNfkeX6e39GHYpQr-g1yU0367OFWeRg");

  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var fieldNames = values[0];//First row of the sheet must be the the field names

  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == appNumber) {
      var row = values[i];
      var body = bodyCopy.copy();
    
      for (var f = 0; f < fieldNames.length; f++) {
        body.replaceText("\\[" + fieldNames[f] + "\\]", row[f]);//replace [fieldName] with the respective data value
      }
    
      var date = Utilities.formatDate(new Date(), "EST", "MM-dd-yyyy");
      body.replaceText("\\[Today\\]", date);
      var numChildren = body.getNumChildren();//number of the contents in the template doc
     
      for (var c = 0; c < numChildren; c++) {//Go over all the content of the template doc, and replicate it for each row of the data.
        var child = body.getChild(c);
        child = child.copy();
        if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
          mergedDoc.appendHorizontalRule(child);
        } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
          mergedDoc.appendImage(child.getBlob());
        } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
          mergedDoc.appendParagraph(child);
        } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
          mergedDoc.appendListItem(child);
        } else if (child.getType() == DocumentApp.ElementType.TABLE) {
          mergedDoc.appendTable(child);
        } else {
          Logger.log("Unknown element type: " + child);
        }
      }
      break;
    }
  };
  var start = findStart(building);
  var directions = Maps.newDirectionFinder()
    .setOrigin(start)
    .setDestination(end)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .getDirections();
      
  // These regular expressions will be used to strip out
  // unneeded HTML tags
  var r1 = new RegExp('<b>', 'g');
  var r2 = new RegExp('</b>', 'g');
  var r3 = new RegExp('<div style="font-size:0.9em">', 'g');
  var r4 = new RegExp('</div>', 'g');
      
  // Much of this code is based on the template referenced in
  // http://googleappsdeveloper.blogspot.com/2010/06/automatically-generate-maps-and.html
  
  for (var j in directions.routes[0].legs) {
    for (var k in directions.routes[0].legs[j].steps) {
      // Parse out the current step in the directions
      var step = directions.routes[0].legs[j].steps[k];
    
      // Pull out the direction information from step.html_instructions
      // Because we only want to display text, we will strip out the
      // HTML tags that are present in the html_instructions
      var text = step.html_instructions;
      text = text.replace(r1, ' ');
      text = text.replace(r2, ' ');
      text = text.replace(r3, ' ');
      text = text.replace(r4, ' ');
      
      // Add each step in the directions to the directionsPanel
      mergedDoc.appendParagraph(text);//Appending page break. Each row will be merged into a new page.
    }
  }
    
  mergedDoc.appendParagraph("");  
  mergedDoc.appendParagraph("Length of trip (miles): " + findDistance(building, end));
  mergedDoc.appendParagraph("Travel time: " + findTime(building, end));
  mergedDoc.saveAndClose();
  return mergedDoc;
}

function findTime (building, end) {
  var start = findStart(building);
    
  var directions = Maps.newDirectionFinder()
    .setOrigin(start)
    .setDestination(end)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .getDirections();
      
  var timeinseconds = directions.routes[0].legs[0].duration.value;
  var timeinminutes = timeinseconds/(60);
  var timeinhours = Math.floor(timeinminutes/60);
  timeinminutes = Math.round(timeinminutes%60);
  
  return timeinhours + " hours and " + timeinminutes + " seconds";
}

function findDistance (building, end) {
  var start = findStart(building);
    
  var directions = Maps.newDirectionFinder()
    .setOrigin(start)
    .setDestination(end)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .getDirections();
      
  var distinmeters = directions.routes[0].legs[0].distance.value;
  var distinmiles = (distinmeters/1609.34).toFixed(2);
  
  return distinmiles;
}

function findStart(building) {
  var start = "26894 Schady Rd, Olmsted Township, OH 44138"
    
  
  if (building == "HS") {
    start = "26939 Bagley Rd, Olmsted Falls, OH 44138";
  } else if (building == "MS") {
    start = "27045 Bagley Rd, Olmsted Falls, OH 44138";
  } else if (building == "IS") {
    start = "27043 Bagley Rd, Olmsted Falls, OH 44138";
  } else if (building == "FL") {
    start = "26450 Bagley Rd, Olmsted Falls, OH 44138";
  } else if (building == "ECC") {
    start = "7105 Fitch Rd, Olmsted Falls, OH 44138";
  }
  
  return start;
}