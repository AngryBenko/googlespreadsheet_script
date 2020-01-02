/*
LAST UPDATED: 1/1/2020
=======================Changelog===========================
- Optimized the code to run faster by reducing the number of reads from the spreadsheet
- Fixed indexing offset background index and row numbers
- Adjusted code for new packet procedure
- Added automatic removal of checkboxes for BFP and RETURNS
- Added thank you card column with automated formatting
- Added a custom menu for formatting the spreadsheet
- Added initial cell value for "Packet" column
- Sort of fixed the spaghetti code. Yay
======================= To-Do =============================

======================= Etc. ==============================

*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Format sheet', 'autoFormat')
      .addSeparator()
      .addSubMenu(ui.createMenu('Hello!')
                  .addSubMenu(ui.createMenu("I don't know why you are here")
                             .addItem("Why are you so curious??", 'nullNothing')
                             )
                  .addSubMenu(ui.createMenu("Seriously, don't go any farther...")
                             .addSubMenu(ui.createMenu("Fo'real tho. If you don't stop, something bad will happen.")
                                         .addSubMenu(ui.createMenu("What do you think this is? Inception?")
                                                     .addSubMenu(ui.createMenu("Sigh..There's just no stopping you is there?")
                                                                 .addSubMenu(ui.createMenu("Alright. Alright. You win...")
                                                                             .addItem("Check back later for a riddle.",'nullNothing')
                                                                            )
                                                                )
                                                    )
                                        )
                             )
                 )
      .addToUi();
  addScriptProp();
  //Utilities.sleep(3000);
  //autoFormat();
}

function nullNothing() {
}

function onEdit() {
  //addScriptProp();
  // Add a wait timer between edits. perhaps 3, 5 secs?
  //Utilities.sleep(3000);
  //autoFormat();
}

// Load script properties
function addScriptProp() {
  var scriptProp = PropertiesService.getScriptProperties();
  scriptProp.setProperties({
    'Sun'  : '#a61c00', // SUNDAY - Maroon/Red
    'Mon'  : '#ff9900', // MONDAY - Orange
    'Tue'  : '#ffff00', // TUESDAY - Yellow
    'Wed'  : '#00ff00', // WEDNESDAY - Green
    'Thu'  : '#00ffff', // THURSDAY - Light Blue
    'Fri'  : '#9900ff', // FRIDAY - Purple
    'Sat'  : '#a64d79',  // SATURDAY - Light Purple
  });
  
}

// Returns 'true' if variable d is a date object.
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function addCheckbox(gSS, chkBoxValues, tYValues, iValue, row_num) {
  var criteria = SpreadsheetApp.DataValidationCriteria.CHECKBOX;
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  
  // j increments through the row
  for(var j = 0; j < chkBoxValues[iValue].length; j++) {
    //Logger.log("THIS  IS ADDCHECKBOX: iValue: " + iValue + " -- j: " + j);
    if(typeof chkBoxValues[iValue][j] == 'string') {
      gSS.getRange('i' + row_num + ':o' + row_num).setDataValidation(rule);
    }
    
    if(typeof tYValues[iValue][0] == 'string') {
      gSS.getRange('s' + row_num + ':s' + row_num).setDataValidation(rule);
    }
  } 
}

function removeCheckbox(gSS, chkBoxValues, thankYouValues, iValue, row_num) {
  for(var j = 0; j < chkBoxValues[iValue].length;j++) {
    gSS.getRange('i' + row_num + ':o' + row_num).setDataValidation(null);
    gSS.getRange('i' + row_num + ':o' + row_num).clearContent();
    gSS.getRange('s' + row_num + ':s' + row_num).setDataValidation(null);
    gSS.getRange('s' + row_num + ':s' + row_num).clearContent();
  }
}

function getBackgroundColorRange(row_num) {
  //Logger.log('b' + 0 + ':q' +  0);
  return 'b' + row_num + ':q' + row_num;
}

function fontBlack(gSS, row_num) {
  gSS.getRange(getBackgroundColorRange(row_num)).setFontColor("#000000");
  gSS.getRange(getBackgroundColorRange(row_num)).setFontColor(null);
}

function fontWhite(gSS, row_num) {
  gSS.getRange(getBackgroundColorRange(row_num)).setFontColor("#ffffff");
}
// Gold #f1c232
function bgGreen(gSS, row_num) {
  gSS.getRange(getBackgroundColorRange(row_num)).setBackground("#d9ead3"); //TODO SET UP AUTOMATIC COLOR FORMATTING, Check for checkboxes(data validation) ----- b : q
  fontBlack(gSS, row_num);
}

function bgPink(gSS, row_num) {
  gSS.getRange(getBackgroundColorRange(row_num)).setBackground("#ffbcf2");
  fontBlack(gSS, row_num);
}

function bgBlack(gSS, row_num) {
  gSS.getRange(getBackgroundColorRange(row_num)).setBackground("#000000");
  fontWhite(gSS, row_num);
}

function bgGrayScale(gSS, row_num) {
  gSS.getRange(getBackgroundColorRange(row_num)).setBackground("#cccccc");
}

function getScriptProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key)
}

function isSpecialColor(curBackground) {
  // Gold, Red, Blue
  if((curBackground != "#f1c232") && (curBackground != "#ea9999") && (curBackground != "#c9daf8")) {
    return true;
  }
  return false;
}

function defaultFormat(ss, rows, rangeEnd) {
  rows.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  rows.setFontFamily("Arial");
  rows.setFontSize(11);
  rows.setHorizontalAlignment("left").setVerticalAlignment("bottom");
  //rows.setFontColor(null);
  
  ss.getRange('i1:o' + rangeEnd).setHorizontalAlignment("center"); // Checkboxes
  ss.getRange('s1:s' + rangeEnd).setHorizontalAlignment("center"); // Thank you
  ss.getRange('t1:t' + rangeEnd).setHorizontalAlignment("center"); // Driver
  ss.getRange('e1:e' + rangeEnd).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Stock
  ss.getRange('c1:c' + rangeEnd).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Model
  ss.getRange('d1:d' + rangeEnd).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Deal #
  ss.getRange('g1:g' + rangeEnd).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // City
  ss.getRange('q1:q' + rangeEnd).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Notes
  // Modify upper, uppercase code to default format for column H(GROUP)
  ss.getRange('a1:t1').setHorizontalAlignment("center").setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK); // Title
 }

function dataCheck(value) {
  Logger.log(value);
  Logger.log(value.length);
}

function formatSheet(curBackground, ss, row_num, stockValues, checkBoxValues, thankYouValues, i, packetValues, bgIndex) {
  if(curBackground[bgIndex] == "#ffffff" || curBackground[bgIndex] == "#f5f7ff") { // If background is white, give it a default color(light green)
    bgGreen(ss, row_num);
  } else { // ELSE, the background is NOT white -> go into conditional
    if(stockValues[i][0] == 'BFP') {
      bgPink(ss, row_num);
      removeCheckbox(ss, checkBoxValues, thankYouValues, i, row_num);
    } else if(stockValues[i][0] == 'RETURN') {
      bgBlack(ss, row_num);
      removeCheckbox(ss, checkBoxValues, thankYouValues, i, row_num);
    } else { // means this is green
      if(isSpecialColor(curBackground[bgIndex])) {
        addCheckbox(ss, checkBoxValues, thankYouValues, i, row_num);
        bgGreen(ss, row_num);
      }
    }
  }
  // If Packet column cell is empty, add initial cell content
  if(packetValues[i][0] == "") {
    SpreadsheetApp.getActiveSheet().getRange('p' + (i+1)).setValue("Printed-\npacket-");
  }
}

function autoFormat() {
  var rangeEnd = 150;
  var ss = SpreadsheetApp.getActiveSheet(); // EDIT TO ONLY APPLY TO CERTAIN SHEET
  var rows = ss.getRange('a1:t' + rangeEnd); // range to apply formatting to ----Will shift to S
  var totalRows = rows.getNumRows(); // no. of rows in the range named above
  var checkBox = ss.getRange('i1:o' + rangeEnd); // i2 causes weird color formatting issues???
  var checkBoxValues = checkBox.getValues(); // ---- shift from i1:o
  var stockValues = ss.getRange('e1:e' + rangeEnd).getValues(); // For BFP or RETURNs
  var date = ss.getRange('a1:a' + rangeEnd);
  var dateValues = date.getValues(); // array of values to be tested (1st column of the range named above)
  var group = ss.getRange('h1:h' + rangeEnd);
  var groupValues = group.getValues();
  var thankYouValues = ss.getRange('s1:s' + rangeEnd).getValues();
  var packetValues = ss.getRange('p1:p' + rangeEnd).getValues();
  
  //=========================//
  var curBackground = []
  var todayDate = new Date();
  todayDate = Utilities.formatDate(todayDate, "GMT-7", "yyyyMMdd");
  //===========================//
  var row_num, curDate, nextDate, dayOfWeek;
  var bgIndex = 0;
  
  defaultFormat(ss, rows, rangeEnd);
  
  // This for loop reads all of the rows background color. This will minimize the amount of reading needed in the script by storing all of the backgrounds in an array.
  // We then perform actions on the array which makes the script feel faster
  for(var i = 0; i <= totalRows - 1; i++) {
    row_num = i + 1;
    curDate = new Date(dateValues[i]); // READ
    if(todayDate <= Utilities.formatDate(curDate, "GMT-7", "yyyyMMdd")) { // Today's Date is <= currentDate, do formatting, else dont do any formatting for previous dates
      curBackground.push(ss.getRange('b' + row_num).getBackground());
    }
  }
  Logger.log(curBackground);
  
  // Starts loop for formatting all rows
  for (var i = 0; i <= totalRows - 1; i++) {
    row_num = i + 1;
    curDate = new Date(dateValues[i]); // READ
    
    if(isValidDate(curDate)) {
      nextDate = new Date(dateValues[i+1]);
      // If dates are different, seperate days with a thick border
      if (Utilities.formatDate(curDate, "GMT-7", "yyyyMMdd") != Utilities.formatDate(nextDate, "GMT-7", "yyyyMMdd") ) {
          ss.getRange('a' + row_num + ':t' + row_num).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK); // format if true ---- a : s
      }
      if(todayDate <= Utilities.formatDate(curDate, "GMT-7", "yyyyMMdd")) { // Today's Date is <= currentDate, do formatting, else dont do any formatting for previous dates
        dayOfWeek = Utilities.formatDate(curDate, "GMT-7", "E");
        
        // Change the color of the dates according to the day of the week
        ss.getRange('a' + row_num).setBackground(getScriptProp(dayOfWeek));
        
        formatSheet(curBackground, ss, row_num, stockValues, checkBoxValues, thankYouValues, i, packetValues, bgIndex);
        bgIndex += 1;
      } else {
        bgGrayScale(ss, row_num);
      }
      
    }
  }
}

function autoBackground() {
  var rangeEnd = 150;
  var ss = SpreadsheetApp.getActiveSheet();
  var rows = ss.getRange('a1:t' + rangeEnd); // range to apply formatting to ----Will shift to S
  var totalRows = rows.getNumRows(); // no. of rows in the range named above
  var checkBox = ss.getRange('i1:o' + rangeEnd); // i2 causes weird color formatting issues???
  var checkBoxValues = checkBox.getValues(); // ---- shift from i1:o
  var date = ss.getRange('a1:a' + rangeEnd);
  var dateValues = date.getValues(); // array of values to be tested (1st column of the range named above)
  var thankYouValues = ss.getRange('s1:s' + rangeEnd).getValues();
  var stockValues = ss.getRange('e1:e' + rangeEnd).getValues(); // For BFP or RETURNs
  var packetValues = ss.getRange('p1:p' + rangeEnd).getValues();
  
  var row_num, curDate, nextDate, dayOfWeek;
  var curBackground = []
  var bgIndex = 0;
  
  var todayDate = new Date();
  todayDate = Utilities.formatDate(todayDate, "GMT-7", "yyyyMMdd");
  
  for(var i = 0; i <= totalRows - 1; i++) {
    row_num = i + 1;
    curDate = new Date(dateValues[i]); // READ
    if(todayDate <= Utilities.formatDate(curDate, "GMT-7", "yyyyMMdd")) { // Today's Date is <= currentDate, do formatting, else dont do any formatting for previous dates
      curBackground.push(ss.getRange('b' + row_num).getBackground());
    }
  }

  //Logger.log(curBackground);

  for(var i = 0; i <= totalRows - 1; i++) {
    row_num = i + 1;
    curDate = new Date(dateValues[i]); // READ
    
    if(isValidDate(curDate)) {
      if(todayDate <= Utilities.formatDate(curDate, "GMT-7", "yyyyMMdd")) { // Today's Date is <= currentDate, do formatting, else dont do any formatting for previous dates
        if(curBackground[bgIndex] == "#ffffff" || curBackground[bgIndex] == "#f5f7ff") { // If background is white, give it a default color(light green)
          bgGreen(ss, row_num);
        } else { // ELSE, the background is NOT white -> go into conditional
          if(stockValues[i][0] == 'BFP') {
            bgPink(ss, row_num);
            removeCheckbox(ss, checkBoxValues, thankYouValues, i, row_num);
          } else if(stockValues[i][0] == 'RETURN') {
            bgBlack(ss, row_num);
            removeCheckbox(ss, checkBoxValues, thankYouValues, i, row_num);
          } else { // means this is green
            if(isSpecialColor(curBackground[bgIndex])) {
              addCheckbox(ss, checkBoxValues, thankYouValues, i, row_num);
              bgGreen(ss, row_num);
            }
          }
        }
        // If Packet column cell is empty, add initial cell content
        if(packetValues[i][0] == "") {
          SpreadsheetApp.getActiveSheet().getRange('p' + (i+1)).setValue("Printed-\npacket-");
        }
        bgIndex += 1;
      } else {
          bgGrayScale(ss, row_num);
      }
    }
  }
}
