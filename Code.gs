/// IsStringNumber checks if a string is a number
function IsStringNumber(possibleNumber) { // Modified from: https://stackoverflow.com/questions/35252684/if-var-isnumber-for-script
    if (!isNaN(parseFloat(possibleNumber)) && isFinite(possibleNumber)) {
        return true;
    } else {
        return false;
    }
}


/// getVideo gets the video object from a YouTube video ID
function getVideo(id) {
  return YouTube.Videos.list('snippet', {'id': id}).items[0];
}


/// AddVideoDescription appends the given description to the given ids exsisting description
function AddVideoDescription(id, chapters, showDescription = true, confirmation = true) {
  var video = getVideo(id);
  var ui = SpreadsheetApp.getUi();


  if (video.snippet.description.indexOf('0:00') > -1) { // Does the video already have chapters in the description?
    var message1 = "Description for video of ID: " + id + " already cotains the text '0:00' are you sure want to continue?\r\n";
    var message2 = "Description is currentely:\r\n\r\n" + video.snippet.description;
    var message3 = "\r\n\r\nDescription will be:\r\n\r\n" + video.snippet.description + '\n\n\n' + chapters;
    var response = ui.alert(message1 + message2 + message3, ui.ButtonSet.YES_NO);
    if(response != ui.Button.YES) {
      ui.alert("Ok, skipping video of ID: " + id + "\r\nConsider setting its column R to TRUE", ui.ButtonSet.OK);
      return false;
    }
  }


  if (video.snippet.description === "") { // if blank, don't add new lines
    video.snippet.description = chapters;
  }
  else { // not blank
    video.snippet.description = video.snippet.description + '\n\n' + chapters; // Append chapters to the existing description
  }
  var newDescription = video.snippet.description;

  // Manual confirmation of Description
  if (showDescription) {
    var response = ui.alert('Does this look right? \r\n\r\n' + newDescription, ui.ButtonSet.YES_NO);
    if (response != ui.Button.YES) { // Cancel
      throw("Canceled, description was not updated"); // Exit program
    }
  }


  YouTube.Videos.update(video, 'snippet'); // Updates the video with the new description


  if (confirmation === true) { // Check to see if video updated correctely
    Utilities.sleep(15000); // Wait 15s, Videos don't update instantly

    video = getVideo(id); // Get video again

    var message1 = "Description changed, is this correct - decriptionCorrect should equal true. \r\ndecriptionCorrect = ";
    var message2 = "\r\n(if description hasn't changed, check actual video description):\r\n\r\n";
    response = ui.alert(message1 + (newDescription === video.snippet.description) + message2 + video.snippet.description, ui.ButtonSet.YES_NO);
    if (response != ui.Button.YES) {
      throw("Exiting, description has still been updated, please fix manually"); // Exit program
    }
  }
  return true;
}


/// oneVideoAddChaptersDescription runs AddVideoDescription for a single user inputed row
function oneVideoAddChaptersDescription(confirmation = true) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet.getName() != "Videos") { // Is user in videos tab?
    throw("Please switch to videos tab");
  }
  

  var inputedRow = Browser.inputBox("Which row do you want to update the description of?", Browser.Buttons.OK_CANCEL);

  if (inputedRow === 'cancel' || inputedRow === '') { // Canceled
    throw("Cancelled");;
  }
  if (inputedRow < 2 || inputedRow > sheet.getMaxRows()) { // To big or to small
    throw("Row: '" + inputedRow + "' out of range");
  }
  if (!IsStringNumber(inputedRow)) { // Is not a number
    throw("input: '" + inputedRow + "' is not a number");
  }


  var url = sheet.getRange("C" + inputedRow).getValue();
  if (!(url.indexOf('https://www.youtube.com/watch?v=') > -1)) { // ensure string contains the start of a link
      throw("C" + inputedRow + " does not contain a valid youtube link");
  }

  var videoID = url.split('https://www.youtube.com/watch?v=')[1]; // just the id now

  var chapters = sheet.getRange("K" + inputedRow).getValue(); // The chapters cell

  if (!(chapters.indexOf('0:00') > -1)) { // Does chapters contain the required 0:00 starting chapter?
    throw("Chapters is missing at least the first chapter ('0:00' not in the chapters text)");
  }

  var success = false;

  if (confirmation) {
    success = AddVideoDescription(videoID, chapters, true, true);
  }
  else {
    success = AddVideoDescription(videoID, chapters, true, false);
  }

  if (success) {
    sheet.getRange("R" + inputedRow).setValue(true); //Now imported
  }
}

/// confirmation = true for oneVideoAddChaptersDescription()
function oneVideoAddChaptersDescriptionT() { //required because of ui.createMenu.addItem limits
  oneVideoAddChaptersDescription(true);
}

/// confirmation = false for oneVideoAddChaptersDescription()
function oneVideoAddChaptersDescriptionF() { //required because of ui.createMenu.addItem limits
  oneVideoAddChaptersDescription(false);
}


/// allVideosAddChaptersDescription traverses an entire sheet and runs AddVideoDescription for every relevant video
function allVideosAddChaptersDescription(debug = false) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getName() != "Videos") { // Is user in videos tab?
    throw("Please switch to videos tab");
  }  

  var confirmMessage1 = "You are about to update every applicable video description."
  var confirmMessage2 = "\r\nDo you want to do this in Debug Mode? (Confirmation before and after every change)"
  var confirm = ui.alert(confirmMessage1 + confirmMessage2, ui.ButtonSet.YES_NO_CANCEL);
  if (confirm == ui.Button.YES) { // Debug on
    debug = true
  }
  else if (confirm != ui.Button.NO) { // Debug is off by default
    throw("Canceled");
  }
  
  var message1 = "Debug Mode is set to " + debug + ", do you want to continue?"
  var message2 = "\r\n\r\nNote: if debug mode is on and you answer no to any question the program will exit."
  confirm = ui.alert(message1 + message2, ui.ButtonSet.YES_NO)
  if (confirm != ui.Button.YES) { // Cancel
    throw("Canceled");
  }

  var totalRows = sheet.getMaxRows(); // Row count
  var totalUpdated = 0; // Each updated video adds one to the count
  var updatedRows = ""; // Append each updated row to the end of this string
  for(var i = 2; i < totalRows+1; i++) { // first video row is row 2
    // not already imported (Row R), is complete (Row Q), and does exists (Row I)
    if(sheet.getRange("R" + i).getValue() == false && sheet.getRange("Q" + i).getValue() == true && sheet.getRange("I" + i).getValue() == true) {
      if(!(sheet.getRange("K" + i).getValue().indexOf('0:00') > -1)) { // no chapters?
        ui.alert("Row " + i + " is set to finished, but it has no chapters written down, skipping", ui.ButtonSet.OK);
      }
      else {
        var url = sheet.getRange("C" + i).getValue();
        if (!(url && url.indexOf('https://www.youtube.com/watch?v=') > -1)) { // ensure value exists and cotains the start of a link
          var message = "C" + i + " doesn't contain a valid youtube link, but it is marked as finished. Please fix this, skipping";
          ui.alert(message, ui.ButtonSet.OK);
        }

        var videoID = url.split('https://www.youtube.com/watch?v=')[1]; // just the id now

        var chapters = sheet.getRange("K" + i).getValue();

        var success = false;
        if (debug) {
          success = AddVideoDescription(videoID, chapters, true, true); // debug
        }
        else {
          success = AddVideoDescription(videoID, chapters, false, false);
        }
        
        if (success) {
          sheet.getRange("R" + i).setValue(true); //Now imported
          updatedRows += i + " ";
          totalUpdated++;
        }
      }
    }
  }

  if (totalUpdated != 0) { // If there were updated rows
    var messageUpdatedRows = "\r\n\r\nFull list of updated rows:\r\n" + updatedRows
    ui.alert("Total rows updated: " + totalUpdated + messageUpdatedRows, ui.ButtonSet.OK);
    return;
  }
  ui.alert("No descriptions updated", ui.ButtonSet.OK);
}



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Add Descriptions')
      .addItem('One Video - Wait for confirmation', 'oneVideoAddChaptersDescriptionT') // a single video - with the 15s wait
      .addItem('One Video - No wait for confirmation', 'oneVideoAddChaptersDescriptionF') // a single video - no wait
      .addItem('All Videos', 'allVideosAddChaptersDescription') // the whole spreadsheet
      .addToUi();
}
