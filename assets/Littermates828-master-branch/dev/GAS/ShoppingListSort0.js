
/******************************************************************************
 * Joe Kelly (JKSFO | Littermates828)
 * 2024-01-27
 *
 * Assumes range A1: S150
 *****************************************************************************/

//0 See if all of these are used ...

// REVIEW FOR ALL TRIGGERS

const vUser = Session.getActiveUser().getEmail();
// const vUser = ""; // = Session.getActiveUser().getEmail();

var gSS = SpreadsheetApp.getActiveSpreadsheet();
var gMain = gSS.getSheets()[0];
var gMainRows = gMain.getDataRange().getValues().length;
var gLists = gSS.getSheets()[1];
var gListsRows = gLists.getDataRange().getValues().length;
var gRowStart = 3, gRowEnd = 150;

var gRange = "A3:S150"; // Playing field
var gSortRow4Value = 1;

function onOpen(){
  var myA1 = "";
  try{
    setDefaults();
  }
  catch (err) {
    console.log(err.Message);
  }
}

// Can't seem to send a fn param from the sheet ...
// Well not actually but function onSelectionChange(e)
// https://developers.google.com/apps-script/guides/triggers#onselectionchangee
// has way too much latency and is called way too much

/*
function onSelectionChange(e) {
  msg ("selChg", 4);
}

function onEdit(e) {
  var foo = e.getValue;
  msg(foo, 4);
  }

function onSelectionChange(e) {
  var foo = e.range.getValue();
  msg("fnS(" + foo + ")", 4);
}

// Test C
// function fnParam (p) {
function fnParam() {
  msg("fnParam(" + p + ")", 4);
  msg("fnParam(sda)", 4);
}

function fnParam2(p) {
  msg("fnParam(" + p + ")", 4); return('ick');
}
*/

function showHiddenCells() {
  msg("showHiddenCells", 4);
}

function testA () {
  ;
  // Call hide empty cell function
  hideEmptyCells();
}

function testB () {
  ;
  // Call show empty cell function
  showEmptyCells();
}

function testC () {
  ;
  // Call show empty cell function
  alert('Do JS alters work?: If so then other stuff should work like writing to local storage.');
  uiHello();
  showEmptyCells();

  // Local storage - setting
  lsTestSet ();
  // Local storage - getting
  var myValTest = "";
  myValTest = localStorage.getItem("mySomeLSKey2");
  msg(myValTest, 4);


//
}

function lsTestSet () {
  localStorage.setItem("mySomeLSKey1", "mySomeLSKey1Value");
  localStorage.setItem("mySomeLSKey2", "mySomeLSKey2Value");
}


function uiHello() {
  // Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
  // user can also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Getting to know you', 'May I know your name?', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    Logger.log('The user\'s name is %s.', response.getResponseText());
  } else if (response.getSelectedButton() == ui.Button.NO) {
    Logger.log('The user didn\'t want to provide a name.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function showEmptyCells () {
  msg("showEmptyCells", 4);
}

function hideEmptyCells () {
  msg("hideEmptyCells", 4);
  ;
}

/*
function mySort (rng, upperLeft, lowerRight, sortUpDown, sortUpDownText) {
  try {
    SpreadsheetApp.getActiveSheet().getRange(rng).sort( {column: lowerRight, ascending: sortUpDown} );
    gMain.getRange(upperLeft, lowerRight).setValue(sortUpDownText);
  }
  catch {
    msg(errText, 4);
  }
}
*/

function mySort(sortValue /* existing value */, col, upperLeft, lowerRight,
  sortUpDown, sortUpDownText /* sort value to be */) {
  try {
    SpreadsheetApp.getActiveSheet().getRange(col).sort({ column: lowerRight, ascending: sortUpDown });
    gMain.getRange(upperLeft, lowerRight).setValue(sortUpDownText);


        var myCell = gMain.getRange(gSortRow4Value, lColumn);
        var mySortAsc = myCell.getValue() ;

        

          // WFT!@#$@#!!$@!$! I hate Google App Script
          if ((mySortAsc == true) | (mySortAsc == "true") |(mySortAsc == "TRUE")) {
            mySort (mySortAsc, gRange, gSortRow4Value, lColumn, false, "FALSE");
          }
          else if ((mySortAsc == false) | (mySortAsc == "false") |(mySortAsc == "FALSE")) {
            mySort (gRange, gSortRow4Value, lColumn, true, "TRUE");
          }
          else{  // Just sort it as ASC
            mySort (gRange, gSortRow4Value, lColumn, false, "FALSE");
            throw("Sort value in first row can be either TRUE or FALSE. /n Defaulted to TRUE or Ascending. ");
          }
          

          
          
          
          
          
          
          
          
          
          
          
          
          
          
          /*
          
          PLAY 
          
          
          onEdit(e) runs when a user changes a value in a spreadsheet.
onSelectionChange(e)

*/
          

/**
 * 
 * CHECK FOR MORE THAN ONE 
 * 
 * The event handler triggered when the selection changes in the spreadsheet.
 * @param {Event} e The onSelectionChange event.
 * @see https://developers.google.com/apps-script/guides/triggers#onselectionchangee
 * /
function onSelectionChange(e) {
  // Set background to red if a single empty cell is selected.
  const range = e.range;
  if (range.getNumRows() === 1 &&
    range.getNumColumns() === 1 &&
    range.getCell(1, 1).getValue() === '') {
    range.setBackground('red');
  }
}          
          
   /*
   
   

Another option is to use the e.parameter.source value to determine the ID of the element that triggered the serverHandler to be called.

Here's an example:

function doGet(e) {
  var app = UiApp.createApplication();
  var handler = app.createServerHandler("buttonAction");

  for (var i = 0; i < 4; i++) {
    app.add(app.createButton('button'+i).setId(i).addClickHandler(handler));
  }
  return app;
}


function buttonAction(e) {
  var app = UiApp.getActiveApplication();
  Logger.log(e.parameter.source);    
}

e.parameter.source will contain the ID of the element, which you could then use to call app.getElementById(e.parameter.source) ...

*/


          
          
          

  }
  catch {
    msg(errText, 4);
  }
}

// SEE IF THIS WORKS AND IS CALLED AS OFTEN ...
//function doGet(e) {}

// NOTE: Vars for cell references don't seem to work so verbose ...
function sortYN(){

  // NOTE: is there a way to get the name and the callee ordinal position?
  // https://stackoverflow.com/questions/10105526/google-apps-script-find-function-caller-id
  // e.parameter.source
  // Change it to a cell.click() event?

  try{
    var lColumn = 1; // CHANGE FOR COLUMN

    // /*
    // Call to ...
    // mySort (sortValue /* existing value */, col, upperLeft, lowerRight,
    // sortUpDown, sortUpDownText /* sort value to be */) {
    // */

    /*
    var myCell = gMain.getRange(gSortRow4Value, lColumn);
    var mySortAsc = myCell.getValue() ;
    */

    /*
    // WFT!@#$@#!!$@!$! I hate Google App Script
    if ((mySortAsc == true) | (mySortAsc == "true") |(mySortAsc == "TRUE")) {
      mySort (mySortAsc, gRange, gSortRow4Value, lColumn, false, "FALSE");
    }
    else if ((mySortAsc == false) | (mySortAsc == "false") |(mySortAsc == "FALSE")) {
      mySort (gRange, gSortRow4Value, lColumn, true, "TRUE");
    }
    else{  // Just sort it as ASC
      mySort (gRange, gSortRow4Value, lColumn, false, "FALSE");
      throw("Sort value in first row can be either TRUE or FALSE. /n Defaulted to TRUE or Ascending. ");
    }
    */
  }
  catch (errText){
    // console.log(errText)
    msg(errText, 4);
  }
}

function sortItems(){ /* Sort by item first, ascending */
try{
  var lColumn = 2; // CHANGE FOR COLUMN
  var myCell = gMain.getRange(gSortRow4Value, lColumn);
  var mySortAsc = myCell.getValue() ;

  // WFT!@#$@#!!$@!$! I hate Google App Script
  if ((mySortAsc == true) | (mySortAsc == "true") |(mySortAsc == "TRUE")) {
    mySort (gRange, gSortRow4Value, lColumn, false, "FALSE");
  }
  else if ((mySortAsc == false) | (mySortAsc == "false") |(mySortAsc == "FALSE")) {
    mySort (gRange, gSortRow4Value, lColumn, true, "TRUE");
  }
  else{  // Just sort it as ASC
    mySort (gRange, gSortRow4Value, lColumn, false, "FALSE");
    throw("Sort value in first row can be either TRUE or FALSE. /n Defaulted to TRUE or Ascending. ");
  }
}
catch (errText){
  // console.log(errText)
  msg(errText, 4);
}
}

function sortGroup(){ /* Group sorted after item, ascending */
  msg("In sortGroup(" + vSortYNAsc + ")", 2);
  /*try{
    if (vSortGroupAsc == true){
      SpreadsheetApp.getActiveSheet().getRange("A2:R150").sort( {column: 3,ascending: vSortGroupAsc} );
      vSortGroupAsc = false;
    }
    else{
      SpreadsheetApp.getActiveSheet().getRange("A2:R150").sort( {column: 3,ascending: vSortGroupAsc} );
      vSortGroupAsc = true;
    }
  }
  catch (err){
    console.log(err.Message)
  }*/
}

function sortDefault(){
  msg(vSortDefaultAsc + "|" + vSortItemAsc + "|" + vSortGroupAsc + "|" + vSortYNAsc , 3);
  sortItem; // (true)
  sortGroup; // (true)
  sortYN; // (true)
}

function iLeftTheCell(){
  msg("iLeftTheCell", 4);
}

function myFunction() {
 // onOpen(e)
 // Logger.log(Session.getActiveUser().getEmail());
}

function msg(what, time){
  SpreadsheetApp.getActive().toast(what, "User: " + vUser, time );
}

function setDefaults(){
  ;
  }


/*
function demo(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange("B2:D4");

  // The row and column here are relative to the range
  // getCell(1,1) in this code returns the cell at B2
  var cell = range.getCell(1, 1);
  Logger.log(cell.getValue());
}

    if (mySortAsc == true) {
      gMain.getRange(5, 5).setValue("true");
    }
    if (mySortAsc == "true") {
      gMain.getRange(6, 6).setValue("/'true'/");
    }
    if (mySortAsc == "TRUE") {
      gMain.getRange(7, 7).setValue("/'TRUE'/");
    }

// https://docs.google.com/spreadsheets/d///18MO8ZwTuoHAUr_U0GeUDKnHWO9VBIEoUlCVliQyNnd0/edit#gid=0&range=H1

*/


