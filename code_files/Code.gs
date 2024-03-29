/////////////////////////////////////////////////////////////
// GLOBAL VARIABLES DEFINED BELOW
//
// Note that linking to items directly by URL (from within the OPAC, for example)
//   can be done with the deployment URL followed by: ?barcode= and then the barcode
//   for example: https://script.google.com/a/macros.../exec?barcode=123456789

// Change FolderID to where scanned documents are located
//   which can be found by selecting "getfolderid" above and click Run
var folderOfScans = '1ROYjGhNoqh9-dVPhKUlo73ePNO4FfTzD'; 

// LOAN LENGTH IN HOURS; now set in the DepedentSelect page.
// var loanLength = 2;

// URL FOR REQUESTING RESERVES: this is the Web App URL (Deployment URL)
var formURL = 'https://script.google.com/a/macros/hcc.edu/s/AKfycbxFPjfiDHGE5MGLhUiXKuRHnLy1-l55kp5oz6qO14Ai_wpJ16FrSIBe7SaPYQ-dLEE7-w/exec';

// SET DOMAIN OF COLLEGE E-MAIL ADDRESSES TO RESTRICT TO ONLY THESE USERS
//   Not needed if this script is deployed with limit of access only to your domain.
//   However, this is useful for informing students who are not logged in that they need to be.
//   Leave blank to disable this limit (blank= '')
var limitToDomain = 'hcc.edu';

// NAME OF SHEET CONTAINING LIST OF ITEMS ON E-RESERVE
//    No change needed unless the sheet is renamed
var shReserves = SpreadsheetApp.getActive().getSheetByName('Reserves'); 
var shSettings = SpreadsheetApp.getActive().getSheetByName('Settings'); 

/////////////////////////////////////////////////////////////
// REST OF SCRIPT

var item_url; // Make this a global variable so accessible to all functions

// GET ID OF FOLDER CONTAINING RESERVES
function getfolderid() { //get ID of folder by name - 7/27/22 JR
  Logger.log(DriveApp.getFoldersByName('CDL Files').next().getId());
}

//-----REQUEST FORM FUNCTIONALITY-------
//DISPLAY THE REQUEST FORM TO PATRONS 
function doGet(e) { 

  // Get User's email address and extract domain.  If it doesn't match "limitToDomain", prompt to login.
  if (limitToDomain !== '') {
    var visitorEmail = Session.getActiveUser().getEmail();
    var domain = visitorEmail.split('@')[1];
    Logger.log("Effective User's email address: " + visitorEmail);
    Logger.log("Domain: " + domain);
      
    // Check if user is logged in; block access until they are
    if (domain !== limitToDomain) {
      var htmlOutput =  HtmlService.createTemplateFromFile('LoginNeeded');
      htmlOutput.formURL = formURL; // Give link to page
      var collegeWebsite = 'https://www.' + limitToDomain;
      htmlOutput.collegeWebsite = collegeWebsite;
      return htmlOutput.evaluate();  
    }
  }
  

  var htmlOutput =  HtmlService.createTemplateFromFile('DependentSelect'); //Build page from DependentSelect.html
  var course = getCourse();
  htmlOutput.message = '';
  htmlOutput.new_table ='';
  htmlOutput.incorrect_barcode_message = '';
  htmlOutput.course = course;

  var barcode = e.parameter.barcode;
  
  // IF BARCODE PARAMETER, SEND THAT HTML PAGE, OTHERWISE SEND USER GENERAL REQUEST PAGE
  if (barcode) {
    // IF BARCODE IS DEFINED, PROCESS SPECIFIC ITEM
    barcode = barcode.toString();

    var barcode_test = checkBarcode(barcode);
    if (barcode_test !== 'yes') {
      // CANNOT FIND TEXTBOOK BASED ON BARCODE SO ASK USER TO SELECT ITEM MANUALLY
      var htmlOutput =  HtmlService.createTemplateFromFile('DependentSelect');
      var course = getCourse();
      htmlOutput.message = '';
      htmlOutput.new_table ='';
      htmlOutput.course = course;
      // Incorrect barcode requested
      var incorrect_barcode_message = 'That barcode, '+barcode+' is incorrect. Please select the reserve item from the menu below.';
      htmlOutput.incorrect_barcode_message = incorrect_barcode_message;
      return htmlOutput.evaluate();
    }

    var item_title = getTitle(barcode); //Grab item barcode
    var itemOutput =  HtmlService.createTemplateFromFile('RequestByBarcode');
    itemOutput.itembarcode = barcode; //Pass on item barcode
    itemOutput.itemtitle = item_title; // Pass on item title
    
    return itemOutput.evaluate(); 
  
  } else {
    return htmlOutput.evaluate();
  }
}


//Get URL of CDL script/web app
function getUrl() { 
 var url = ScriptApp.getService().getUrl();
 return url;
 Logger.log(url);
}


//Patron form submission 
function doPost(e) {
  Logger.log(JSON.stringify(e)); //Log transaction 
  var name = e.parameters.name.toString(); //Convert name to string
  var studentid = e.parameters.studentid.toString(); //Convert student ID to string

  // COULD GET E-MAIL FOR USER IF LIBRARY ISN'T A GOOGLE SUITE SCHOOL
  // var email = e.parameters.email.toString(); //Convert student ID to string
  var email = Session.getActiveUser().getEmail();  // Log the email address of the person running the script.
  const now = new Date(); //Create a date object for the current date and time 

  var item_url;
  var array_temp = [];
  var barcode = e.parameter.barcode;

  if (barcode) {
    // BARCODE FOUND IN POST SO TREAT AS A DIRECT REQUEST FROM OPAC
    barcode = barcode.toString();

    var barcode_test = checkBarcode(barcode);
    if (barcode_test !== 'yes') {
      // CANNOT FIND TEXTBOOK BASED ON BARCODE SO ASK USER TO SELECT ITEM MANUALLY
      var htmlOutput =  HtmlService.createTemplateFromFile('DependentSelect'); 
      var course = getCourse();
      htmlOutput.message = '';
      htmlOutput.new_table ='';
      htmlOutput.course = course;
      // Incorrect barcode requested
      var incorrect_barcode_message = 'That barcode, '+barcode+', is incorrect. Please select the reserve item from the menu below.';
      htmlOutput.incorrect_barcode_message = incorrect_barcode_message;
      return htmlOutput.evaluate();
    }

    var course = 'direct_request';
    var textbook;
    array_temp = getUrlAndId(barcode); //Store getBarcodeAndUrl function output in an array
    textbook = array_temp[0]; //Grab item barcode
    item_id = array_temp[1]; //Grab item ID -- Added 6/7/22 EL 
    item_url = array_temp[2]; //Grab item (PDF) URL

  } else {
    // NO BARCODE FOUND IN POST SO TREAT AS A COURSE REQUEST
    var course = e.parameters.course.toString(); //Convert course name to string
    var textbook = e.parameters.textbook.toString(); //Convert textbook (PDF) name to string
    array_temp = getBarcodeAndUrlAndId(course, textbook); //Store getBarcodeAndUrl function output in an array
    barcode = array_temp[0]; //Grab item barcode
    item_id = array_temp[1]; //Grab item ID -- Added 6/7/22 EL 
    item_url = array_temp[2]; //Grab item (PDF) URL  
  }

  /* SET LOAN DATE */
  var date_loan;
  var loanLength = parseFloat(e.parameters.loan); //Get loan duration from dropdown
  if (!((loanLength < 301)&&(loanLength > 0))) { //  Check to see value is reasonable: 1 to 300 minutes, otherwise set to 1 hour
    loanLength = 60;
  }

  date_loan = getLoanDate(barcode,textbook,loanLength); //Store getLoanDate function output in an array
  var date_lend = date_loan[0]; //Grab loan date
  var date_expire = date_loan[1]; //Grab expiration date 
  var loan_status = date_loan[2];  //Grab loan status - Added 6/6/22 EL
  var request_status = date_loan[3]; //Grab request status message - Added 6/17/22 JR
  
  AddRecord(name, studentid, course, textbook, email, barcode, item_id, item_url, loan_status, date_lend, date_expire); //Call AddRecord function
  FileShare(email, item_id, loan_status, date_expire,loanLength); //Call FileShare function 
  var item_table = createTable(course,textbook,date_expire,loan_status,loanLength); //create table for items - added 6/29/22 JR

  if (loan_status === 'In Use') {
    var htmlOutput =  HtmlService.createTemplateFromFile('ResponseInUse');
    var course = getCourse();
    htmlOutput.message = request_status; //Print request status message to web page for patron - modified 6/17/22 JR
    htmlOutput.course = course; //Display courses to patrons in dropdown
    htmlOutput.new_table = item_table; //Display items table - added 6/29/22 JR
    htmlOutput.incorrect_barcode_message = '';
    return htmlOutput.evaluate();  

} else if (loan_status === 'Active') {
    var htmlOutput =  HtmlService.createTemplateFromFile('ResponseSuccess');
    var course = getCourse();
    htmlOutput.message = request_status; //Print request status message to web page for patron - modified 6/17/22 JR
    htmlOutput.course = course; //Display courses to patrons in dropdown
    htmlOutput.new_table = item_table; //Display items table - added 6/29/22 JR
    htmlOutput.item_url = item_url; // Give direct link to item
    htmlOutput.incorrect_barcode_message = '';
    return htmlOutput.evaluate();  
      
  } else {
    var htmlOutput =  HtmlService.createTemplateFromFile('DependentSelect');
    var course = getCourse();
    htmlOutput.message = request_status; //Print request status message to web page for patron - modified 6/17/22 JR
    htmlOutput.course = course; //Display courses to patrons in dropdown
    htmlOutput.new_table = item_table; //Display items table - added 6/29/22 JR
    htmlOutput.incorrect_barcode_message = '';
    return htmlOutput.evaluate();  
  }

}


//DISPLAY UNAVAILABLE ITEMS IN A TABLE WITH NEXT DATE AVAILABLE - added 6/29/22 JR
function createTable(course,textbook,date_expire,loan_status,loanLength) {
  if (loan_status === 'In Use') { //create table for unavailable items - added 6/29/22 JR
    const options = {weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'}
    var day = date_expire.toLocaleDateString(undefined, options);
    var time = date_expire.toLocaleTimeString();
    var new_table = '<th scope="col">Course</th><th scope="col">Title</th><th scope="col">In Use Until</th><tr><td>' + course +'</td><td>'+ textbook +'</td><td>' + day +" at "+ time +'</td></tr>';
  }
  
  else if (loan_status === 'Active') {//CREATE TABLE FOR LOANED ITEM WITH EXPIRATION DATE
    const options = {weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'}
    var date_expire = new Date(date_expire);
    var day = date_expire.toLocaleDateString(undefined, options);
    var time = date_expire.toLocaleTimeString();
    var new_table = '<th scope="col">Course</th><th scope="col">Title</th><th scope="col">Expires</th><tr><td>' + course +'</td><td>'+ textbook +'</td><td>'+ day + " at "+ time +'</td></tr>';
  } 
  else {
    var new_table = '';
  }
  return new_table;
 
  }

//--------GRABBING ALL OF THE VARIABLES --------
// GET COURSE NAMES FROM "RESERVES" SPREADSHEET
function getCourse() { 
  var getLastRow = shReserves.getLastRow();
  var return_array = [];
  for(var i = 2; i <= getLastRow; i++) //ITERATE THROUGH THE ROWS OF THE "RESERVES" SPREADSHEET
  {
      if(return_array.indexOf(shReserves.getRange(i, 1).getValue()) === -1) {
        return_array.push(shReserves.getRange(i, 1).getValue()); //Grab course names from row 1 
      }
  }
  return return_array;  
}

// GET NAMES OF ITEMS FOR A GIVEN COURSE FROM THE "RESERVES" SPREADSHEET
function getTextbook(course) { 
  var getLastRow = shReserves.getLastRow();
  var return_array = [];
  for(var i = 2; i <= getLastRow; i++) //Iterate through rows of the "Reserves" spreadsheet
  {
    if(shReserves.getRange(i, 1).getValue() === course) { //If row 1 matches selected course name 
      return_array.push(shReserves.getRange(i, 2).getValue()); //Grab textbook name(s) from row 2 
    }
  }
  return return_array;  
}

// GET ITEM BARCODE AND PDF URL FROM COURSE AND TEXTBOOK 
function getBarcodeAndUrlAndId(course, textbook) { 
  var getLastRow = shReserves.getLastRow();
  var return_array = [];
  for(var i = 2; i <= getLastRow; i++) //Iterate through rows of the "Reserves" spreadsheet 
  {
    if(shReserves.getRange(i, 1).getValue() === course) { //If row one contains course name selected from dropdown
      if(shReserves.getRange(i, 2).getValue() === textbook) { //If row two contains textbook name selected from dropdown
        return_array.push(barcode_temp = shReserves.getRange(i, 3).getValue().toString()); //Grab item barcode from row 3
        return_array.push(id_temp = shReserves.getRange(i, 4).getValue()); //Grab item ID from row 4 -- Added 6/7/22 EL 
        return_array.push(url_temp = shReserves.getRange(i, 5).getValue()); //Grab PDF URL from row 5 
      }
        
    }
  }
  return return_array; 
}

// Get item textbook, ID, and URL from barcode
function getUrlAndId(barcode) { 
  var getLastRow = shReserves.getLastRow();
  var return_array = [];
  for(var i = 2; i <= getLastRow; i++) //Iterate through rows of the "Reserves" spreadsheet 
  {
      if(shReserves.getRange(i, 3).getValue().toString() === barcode) { //If row one contains barcode
          return_array.push(textbook_temp = shReserves.getRange(i, 2).getValue()); //Grab textbook / title
          return_array.push(id_temp = shReserves.getRange(i, 4).getValue()); //Grab item ID from row 4 -- Added 6/7/22 EL 
          return_array.push(url_temp = shReserves.getRange(i, 5).getValue()); //Grab PDF URL from row 5 
      }
  }
  return return_array; 
}


// Get loan & expiration dates for given item barcode - updated to add textbook 7/29/22 JR
function getLoanDate(barcode, textbook, loanLength) {
  var ssInUse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InUse"); // Grab "InUse" spreadsheet
  var getLastRow = ssInUse.getLastRow();
  var return_array = [];
  
  for(var i = 2; i <= getLastRow; i++) // Iterate through rows of the "InUse" spreadsheet 
  {
    //REQUESTED ITEM IN USE
      if(ssInUse.getRange(i, 1).getValue().toString() === barcode) { // If item barcode is in row one
        var loan_status = 'In Use'; // Set loan status to "In Use"
        var date_lend = ssInUse.getRange(i, 3).getValue() //Set loan date to previous loan's exp date for the email message - updated 8/2/22 EL
        var request_status = 'Item in use. Try again later.'; //store request status message for printing - added 6/17/22 JR
        return_array.push(''); // Get blank lend date - updated 6/28/22 JR
        return_array.push(date_lend); // Store expiration date in array - updated 8/2/2022 EL 
        return_array.push(loan_status); // Store loan status in array
        return_array.push(request_status); //Store request status message in array - added 6/17/22 JR
      }
  }
  //REQUESTED ITEM AVAILABLE 
  if (return_array.length < 1) { // If barcode not found in "InUse" spreadsheet 
    var loan_status = 'Active'; // Give request status "Active" - Added 6/6/22 EL
    var request_status = 'Request Successful. Check your email.'; //store request status message for printing - added 6/17/22 JR
    const now = new Date(); // Create a date object for the current date and time 
    let date_exp = new Date(); // Copy current date object 

    date_exp.setMinutes(date_exp.getMinutes() + loanLength); // Set expiration date-time to loanLength from drop down
    const date_lend = Utilities.formatDate(now, 'America/New_York', 'M/dd/yyyy HH:mm:ss'); // Format lend date
    date_exp = Utilities.formatDate(date_exp, 'America/New_York', 'M/dd/yyyy HH:mm:ss'); // Format expiration date 

 ssInUse.appendRow([barcode, textbook, date_exp]); //Add row to InUse spreadsheet - Added 6/6/22 EL - updated 7/27/22 JR
    return_array.push(date_lend); // Store loan date in array
    return_array.push(date_exp); // Store expiration date in array 
    return_array.push(loan_status); // Store loan status in array 
    return_array.push(request_status); //Store request status message in array - added 6/17/22 JR
  }
  return return_array;  
}

//Add request to "Transactions spreadsheet"
function AddRecord(name, studentid, course, textbook, email, barcode, item_id, item_url, loan_status, date_lend, date_expire) { // Grab variables collected from form and functions
  if(loan_status == 'In Use') { // Do not add expiration date if there is none - 10/14/22 CT
    date_expire = '';
  }
  var ssTransactions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions"); // Get "Transactions" spreadsheet
  ssTransactions.appendRow([name, studentid, course, textbook, new Date(), email, barcode, item_id, item_url, loan_status, date_lend, date_expire]); // Add a row to "Transactions" spreadsheet with values
  
}

// GET NAME OF ITEM FROM A GIVEN BARCODE FROM THE "RESERVES" SPREADSHEET
function getTitle(itemBarcode) { 
  var getLastRow = shReserves.getLastRow();
  var item_title;
  for(var i = 2; i <= getLastRow; i++) // ITERATE THROUGH ROWS OF THE "RESERVES" SPREADSHEET
  {
      if(shReserves.getRange(i, 3).getValue().toString() === itemBarcode) { //If column 3 matches barcode
        item_title = shReserves.getRange(i, 2).getValue(); //Grab textbook name(s) from row 2 
      }
  }
  return item_title;  
}

// Check for valid barcode
function checkBarcode(barcode) { 
  var getLastRow = shReserves.getLastRow();
  var valid_barcode;
  for(var i = 2; i <= getLastRow; i++) //Iterate through rows of the "Reserves" spreadsheet
  {
      if(shReserves.getRange(i, 3).getValue().toString() === barcode) { //If column 3 matches barcode
        valid_barcode = 'yes';
      }
  }
  return valid_barcode;  
}

// Get barcode of item from a given item_id from the "Reserves" spreadsheet
function getBarcode(item_id) { 
  var getLastRow = shReserves.getLastRow();
  var itemBarcode;
  for(var i = 2; i <= getLastRow; i++) //Iterate through rows of the "Reserves" spreadsheet
  {
      if(shReserves.getRange(i, 4).getValue().toString() === item_id) { //If column 4 matches item_id
        itemBarcode = shReserves.getRange(i, 3).getValue(); //Grab barcode from column 3
        Logger.log(itemBarcode + " ID " + item_id)
      }
  }
  return itemBarcode;  
}

// IMPORT CSS TO DependentSelect.html - added 7/8/22 JR
function include(filename) {
  return HtmlService.createHtmlOutputFromFile('css')
      .getContent();
}


//---------SHARE FILE--------
function FileShare(email, item_id, loan_status, date_expire, loanLength) {
  if(loan_status === 'Active'){
    try{
      if (loanLength<60) {
        var loanMessage = loanLength+" minutes";
      } else {
        var loanHours = loanLength / 60;
        var loanMessage = loanHours+" hour(s)";
      }

      var customMessage = "This loan will expire in "+loanMessage+". Please re-request this title if you need more time.  Here is the link to it: "+item_url;  // Please set the custom message here.

      var resource = {role: "reader", type: "user", value: email};
      Drive.Permissions.insert(resource, item_id, {emailMessage: customMessage});

      // The above may not worked for a shared google drive.  If it does not, try using the following instead. 
      //    The important change is using: supportsAllDrives
      //
      //var optionalArgs = {
      //  sendNotificationEmails: true,
      //  supportsAllDrives: true,
      //  emailMessage: customMessage
      //}
      // Drive.Permissions.insert(resource, item_id, optionalArgs);


     // Send email notification; add this notification if the above stops working
     // MailApp.sendEmail(email, 'Item shared with you', customMessage)

      }
    catch (e) {
      Logger.log(item_id);
      Logger.log(e);
      }
  }
  else{
    date_expire.setMinutes(date_expire.getMinutes() + 5);
    var dateExpireClean = Utilities.formatDate(date_expire, "America/New_York", "hh:mm a");
    var itemBarcode = getBarcode(item_id);
    MailApp.sendEmail(email,'Requested Reserve Item Unavailable','The reserve item you requested is currently checked out and will next be available at ' + dateExpireClean + '. To re-request this item please return to our e-reserves request form: ' + formURL + '/exec?barcode=' + itemBarcode); 
   
  }
}

//---------UNSHARE FILE----------
function FileUnshare() {
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var ssTransactions = ss.getSheetByName("Transactions");
  var ssInUse = ss.getSheetByName("InUse")
  var getLastRowTransactions = ssTransactions.getLastRow();
  var getLastRowInUse = ssInUse.getLastRow();
  
  for(var i = 2; i <= getLastRowTransactions; i++) //Iterate through rows of the "Transactions" spreadsheet 
  {
      if(ssTransactions.getRange(i, 10).getValue() === 'Active') { // If request is "Active"
        email = ssTransactions.getRange(i, 6).getValue(); //Get patron email
        item_id = ssTransactions.getRange(i, 8).getValue(); // Get item ID
        date_exp = ssTransactions.getRange(i, 12).getValue(); // Get expiration date 
         
        const now = new Date(); //Create a date object for the current date and time 
        
        if(date_exp === now || date_exp < now ){
          try{
            ssTransactions.getRange(i,10).setValue('Expired');
            asset = DriveApp.getFileById(item_id) ? DriveApp.getFileById(item_id) : DriveApp.getFolderById(item_id);
            asset.removeViewer(email);
            MailApp.sendEmail(email, 'Item Expired', 'Your virtual loan has expired. If you would like to borrow this document for additional time, please place a new request.')
            }
          catch (e) {
            Logger.log("File could not be unshared, please check if email address is valid");
            }
          }
        }
      }
  for(var i = 2; i <= getLastRowInUse; i++){
    const now = new Date();
    date_exp_in_use = ssInUse.getRange(i, 3).getValue();
    Logger.log(date_exp_in_use);
    if(date_exp_in_use === now || date_exp_in_use < now){
      ssInUse.deleteRow(i);
      }
      }
    }


// LIST FILES AND FOLDERS AND PUSH TO SPREADSHEET
function listFilesAndFolders() {
  
  shReserves.clear();
  shReserves.appendRow(["folder", "name", "description","ID", "URL", "size", "update"]);
  try {
    var parentFolder =DriveApp.getFolderById(folderOfScans);
    listFiles(parentFolder,parentFolder.getName())
    listSubFolders(parentFolder,parentFolder.getName());
  } catch (e) {
    Logger.log(e.toString());
  }
}

function listSubFolders(parentFolder,parent) {
  var childFolders = parentFolder.getFolders();
  while (childFolders.hasNext()) {
    var childFolder = childFolders.next();
    Logger.log("Fold : " + childFolder.getName());
    listFiles(childFolder,parent)
    listSubFolders(childFolder,parent + "|" + childFolder.getName());
  }
}

function listFiles(fold,parent){
  var data = [];
  var files = fold.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    
    data = [ 
      fold.getName(),
      file.getName(),
      file.getDescription(),
      file.getId(),
      file.getUrl(),
      file.getSize(),
      file.getLastUpdated()
      ];
    shReserves.appendRow(data);
    shReserves.sort(2).sort(1);
  }
}// END OF LIST FILES AND FOLDERS AND PUSH TO SPREADSHEET

