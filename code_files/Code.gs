/////////////////////////////////////////////////////////////
// GLOBAL VARIABLES DEFINED BELOW
//
// Note that linking to items directly by URL (from within the OPAC, for example)
//   can be done with the deployment URL followed by: ?barcode= and then the barcode
//   for example: https://script.google.com/a/macros.../exec?barcode=123456789

//Change folderName to folder where scanned documents are located before running getfolderid
var folderName = ' '

// GET ID OF FOLDER CONTAINING RESERVES
function getfolderid() { //get ID of folder by name 
  Logger.log(DriveApp.getFoldersByName(folderName).next().getId());
}

// Change FolderID to where scanned documents are located
//   which can be found by selecting "getfolderid" above and click Run
var folderOfScans = ' '; 

// DEFAULT LOAN LENGTH DEFAULT IN MINUTES
var loanDefault = 120; 

// URL FOR REQUESTING RESERVES: this is the Web App URL (Deployment URL)
var formURL = ' ';

// SET DOMAIN OF COLLEGE E-MAIL ADDRESSES TO RESTRICT TO ONLY THESE USERS
//   Not needed if this script is deployed with limit of access only to your domain.
//   However, this is useful for informing students who are not logged in that they need to be.
//   Leave blank to disable this limit (blank= '')
var limitToDomain = ' ';

// NAME OF SHEET CONTAINING LIST OF ITEMS ON E-RESERVE
//    No change needed unless the sheet is renamed
var shReserves = SpreadsheetApp.getActive().getSheetByName('Reserves'); 
// NAME OF SHEET CONTAINING LIST OF ITEMS IN USE
//    No change needed unless the sheet is renamed
var ssInUse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InUse"); 
//NAME OF SHEET CONTAINING LIST OF TRANSACTIONS
// no change needed unless sheet is renamed
var ssTransactions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions"); 

//HTML TEMPLATE NAMES
var dependentSelect = 'DependentSelect'; //name of main CDL request form template
var requestByBarcode = 'RequestByBarcode'; //name of direct barcode request form template
var loanReturn = 'ReturnEarly'; //name of loan return form template
var loginNeeded = 'LoginNeeded'; //name of login prompt template


/////////////////////////////////////////////////////////////
// REST OF SCRIPT

var item_url; // Make this a global variable so accessible to all functions

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
  var returnEarly = e.parameter.returnearly; //if return early url used

  //IF RETURNEARLY LINK USED SEND THE RETURNEARLY HTML PAGE, IF NOT CHECK FOR BARCODE PARAMETER
  if (returnEarly) { 
    var htmlOutput = HtmlService.createTemplateFromFile(loanReturn);
    var email = Session.getActiveUser().getEmail();
    var activeRequest = getActive(email); //get active requests for logged in user 
    if(activeRequest.length > 0) { //if patron has active loans
    //--CUSTOM MESSAGE OPTION-- 
    htmlOutput.message =  email + ' has ' + activeRequest.length + ' loans'
    htmlOutput.activeRequest = activeRequest //pass on active request titles
    }
    else { //if patron has no active loans
      htmlOutput.activeRequest = activeRequest //pass on active request titles
      //--CUSTOM MESSAGE OPTION-- 
      htmlOutput.message = 'No active loans for ' + email
    }
    return htmlOutput.evaluate();
  } 
  
  // IF BARCODE PARAMETER, SEND THAT HTML PAGE, OTHERWISE SEND USER GENERAL REQUEST PAGE
  else if (barcode) {
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
      //--CUSTOM MESAGE OPTION--
      var incorrect_barcode_message = 'The item with barcode '+barcode+' is not available. Please select the reserve item from the menu above.';
      htmlOutput.message = incorrect_barcode_message;
      return htmlOutput.evaluate();
    }

    var item_title = getTitle(barcode); //Grab item barcode -- NEW
    var itemOutput =  HtmlService.createTemplateFromFile(requestByBarcode);
    itemOutput.message = '';
    var textbook = getTitlebyBarcode(barcode); //find all titles with the reuqested barcode as an array 
    itemOutput.textbook = textbook; //Pass on textbook titles
    itemOutput.itembarcode = barcode; //Pass on item barcode
    itemOutput.itemtitle = item_title; // Pass on item title -- NEW
    
    return itemOutput.evaluate(); 
  
  } else {
    return htmlOutput.evaluate();
  }
}

//import css to html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile('css')
      .getContent();
}

//Get URL of CDL script/web app
function getUrl() { 
 var url = ScriptApp.getService().getUrl();
 return url;
}

//PATRON FORM SUBMISSION
function doPost(e) {

  Logger.log(JSON.stringify(e)); //Log transaction 
  const now = new Date(); //Create a date object for the current date and time 
  var item_url;
  var array_temp = [];
  var barcode = e.parameter.barcode;
  var returnEarly = e.parameter.returnEarly;

  // COULD GET E-MAIL FOR USER IF LIBRARY ISN'T A GOOGLE SUITE SCHOOL
  // var email = e.parameters.email.toString(); //Convert student ID to string
  var email = Session.getActiveUser().getEmail();  // Log the email address of the person running the script.

  //RETURN EARLY TITLE FOUND IN POST SO COMPLETE A RETURN
  if(returnEarly) { 
    var htmlOutput = HtmlService.createTemplateFromFile(loanReturn);
    try {
      returnLoan(email,returnEarly); //update expiration date on transactions ss and in use ss to now
      FileUnshare();//unshare expired files
      //CUSTOM MESSAGE OPTION
      htmlOutput.message = returnEarly +' returned!';
      }
    catch (e) {
      Logger.log(email);
      Logger.log(returnEarly);
      Logger.log(e);
      //CUSTOM MESSAGE OPTION
      htmlOutput.message = returnEarly +' return failed';
    }
      
    var activeRequest = getActive(email); //get active requests for logged in user 
    htmlOutput.activeRequest = activeRequest;
    return htmlOutput.evaluate();

  // BARCODE FOUND IN POST SO TREAT AS A DIRECT REQUEST FROM OPAC
  
  } else if (barcode) {
    barcode = barcode.toString();
    var barcode_test = checkBarcode(barcode);
    if (barcode_test !== 'yes') {
      // CANNOT FIND TEXTBOOK BASED ON BARCODE SO ASK USER TO SELECT ITEM MANUALLY
      var htmlOutput =  HtmlService.createTemplateFromFile('DependentSelect'); 
      var course = getCourse();
      htmlOutput.message = '';
      htmlOutput.new_table ='';
      htmlOutput.course = course;
      //CUSTOM MESSAGE OPTION
      var incorrect_barcode_message = 'The item with barcode '+barcode+' is not available. Please select the reserve item from the menu below.';
      htmlOutput.message = incorrect_barcode_message;
      return htmlOutput.evaluate();
    }
  
    var name = e.parameters.name.toString(); //Convert name to string 
    var studentid = e.parameters.studentid.toString(); //Convert student ID to string
    var course = getCourseByBarcode(barcode);
    var textbook = e.parameters.textbook.toString(); //get name of item from dropdown selection
    array_temp = getUrlAndId(textbook); //Store getBarcodeAndUrl function output in an array
    textbook = array_temp[0]; //Grab title of item to share
    item_id = array_temp[1]; //Grab item ID
    item_url = array_temp[2]; //Grab item (PDF) URL

  } else {
    // NO BARCODE FOUND IN POST SO TREAT AS A COURSE REQUEST
    var name = e.parameters.name.toString(); //Convert name to string 
    var studentid = e.parameters.studentid.toString(); //Convert student ID to string
    var course = e.parameters.course.toString(); //Convert course name to string
    var textbook = e.parameters.textbook.toString(); //Convert textbook (PDF) name to string
    array_temp = getBarcodeAndUrlAndId(course, textbook); //Store getBarcodeAndUrl function output in an array
    barcode = array_temp[0]; //Grab item barcode
    item_id = array_temp[1]; //Grab item ID 
    item_url = array_temp[2]; //Grab item (PDF) URL  
  }

  /* SET LOAN DATE */
    var loanLength = parseFloat(e.parameters.loan); //Get loan duration from dropdown
    if (!((loanLength < 301)&&(loanLength > 0))) { //  Check to see value is reasonable: 1 to 300 minutes, otherwise set to loan default global variable
    loanLength = loanDefault;
  }

  date_loan = getLoanDate(barcode,textbook,loanLength); //Store getLoanDate function output in an array
  var date_lend = date_loan[0]; //Grab loan date
  var date_expire = date_loan[1]; //Grab expiration date 
  var loan_status = date_loan[2];  //Grab loan status
  var request_status = date_loan[3]; //Grab request status message 
  
  AddRecord(name, studentid, course, textbook, email, barcode, item_id, item_url, loan_status, date_lend, date_expire); //Call AddRecord function
  FileShare(email, item_id, loan_status, date_expire,loanLength); //Call FileShare function 
  var item_table = createTable(course,textbook,date_expire,loan_status,loanLength); //create table for items 

if (loan_status === 'In Use') {
    var htmlOutput =  HtmlService.createTemplateFromFile('ResponseInUse');
    var course = getCourse();
    htmlOutput.message = request_status; //Print request status message to web page for patron 
    htmlOutput.course = course; //Display courses to patrons in dropdown
    htmlOutput.new_table = item_table; //Display items table - added 6/29/22 JR
    htmlOutput.incorrect_barcode_message = '';
    return htmlOutput.evaluate();  

} else if (loan_status === 'Active') {
    var htmlOutput =  HtmlService.createTemplateFromFile('ResponseSuccess');
    var course = getCourse();
    htmlOutput.message = request_status; //Print request status message to web page for patron 
    htmlOutput.course = course; //Display courses to patrons in dropdown
    htmlOutput.new_table = item_table; //Display items table - added 6/29/22 JR
    htmlOutput.item_url = item_url; // Give direct link to item
    htmlOutput.incorrect_barcode_message = '';
    return htmlOutput.evaluate();  
      
  } else {
    var htmlOutput =  HtmlService.createTemplateFromFile('DependentSelect');
    var course = getCourse();
    htmlOutput.message = request_status; //Print request status message to web page for patron 
    htmlOutput.course = course; //Display courses to patrons in dropdown
    htmlOutput.new_table = item_table; //Display items table - added 6/29/22 JR
    htmlOutput.incorrect_barcode_message = '';
    return htmlOutput.evaluate();  
  }

}

//Display unavailable items in a table with next date available 
  function createTable(course,textbook,date_expire,loan_status) {
    if (loan_status === 'In Use') { //create table for unavailable items 
      const options = {weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'}
      //date_expire.setMinutes(date_expire.getMinutes() + 5);
      var day = date_expire.toLocaleDateString(undefined, options);
      var time = date_expire.toLocaleTimeString();
      var new_table = '<th scope="col">Course</th><th scope="col">Title</th><th scope="col">In Use Until</th><tr><td>' + course +'</td><td>'+ textbook +'</td><td>' + day +" at "+ time +'</td></tr>';
    }
  
  else if (loan_status === 'Active') {//create table for loaned item with expiration date
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

  //Get requests with "active" status for logged in user
  function getActive(email) { 
    var getLastRow = ssTransactions.getLastRow() - 1; //get length of spreadsheet
    if (getLastRow < 1) { //if Transactions sheet is empty rows to get = 1
      var getLastRow = 1;
    }
    var return_array = [];
    var transactions_array = ssTransactions.getRange(2,1,getLastRow,12).getValues(); //get values from the first 12 columns of the transactions spread sheet

    for(var i = 0; i <= getLastRow - 1; i++) { 
      if(transactions_array[i]['9'].toString() === 'Active' && transactions_array[i]['5'].toString()=== email) { //grab active requests
        return_array.push(transactions_array[i]);
      }
    } 
    return return_array;
  }

  //Return loan early by changing the expiration date
  function returnLoan(email,returnEarly) { 
    const now = new Date();
    var getLastRowTransactions = ssTransactions.getLastRow() - 1; //get length of Transactions spreadsheet
    if (getLastRowTransactions < 1) { //if Transactions sheet is empty rows to get = 1
      var getLastRowTransactions = 1;
    }
    var transactions_array = ssTransactions.getRange(2,1,getLastRowTransactions,12).getValues(); //get values from the first 12 columns of the transactions spread sheet
    var getLastRowInUse = ssInUse.getLastRow() - 1; //get length of In Use spreadsheet
    if (getLastRowInUse < 1) { //if In Use sheet is empty rows to get = 1
      var getLastRowInUse = 1;
    }
    var inuse_array = ssInUse.getRange(2,1,getLastRowInUse,3).getValues(); //get first 3 columns of In Use ss
    //change expiration date on transactions ss and in use ss
    for(var i = 0; i <= getLastRowTransactions - 1; i++) { // Iterate through rows of the "Transactions" array 
      if(transactions_array[i]['9'].toString() === 'Active' && transactions_array[i]['5'].toString()=== email && transactions_array[i]['3'] === returnEarly) { //find transaction for loan selected by user
        ssTransactions.getRange(i+2,12).setValue(now);
      }
    }
      for(var i = 0; i <= getLastRowInUse - 1; i++) { // Iterate through rows of the "InUse" array 
      if(inuse_array[i][1].toString() === returnEarly) {
        ssInUse.getRange(i+2,3).setValue(now);
      }
      }
  }

  //Get course names from "Reserves" spreadsheet
  function getCourse() {
    var getLastRow = shReserves.getLastRow() - 1; //get length of spreadsheet
    var return_array = [];
    var reserves_array = shReserves.getRange(2,1,getLastRow,1).getValues(); //get values from the "Reserves" spreadsheet
    for(var i = 0; i <= getLastRow - 1; i++) //Iterate through the reserves_array
    {
       
      if (return_array.indexOf(reserves_array[i].toString()) === -1) { //Grab course names from row 1 
        return_array.push(reserves_array[i].toString());
      }
    }
    return return_array;
  }

// GET NAMES OF ITEMS FOR A GIVEN COURSE FROM THE "RESERVES" SPREADSHEET
 function getTextbook(course) {
    var getLastRow = shReserves.getLastRow() - 1; //get length of spreadsheet
    var return_array = [];
    var reserves_array = shReserves.getRange(2,1,getLastRow,3).getValues();
    for(var i = 0; i <= getLastRow - 1; i++) //Iterate through the reserves_array
    {
        if(reserves_array[i]['0'].toString()=== course && return_array.indexOf(reserves_array[i]['1'].toString()) === -1) { //If column 1 matches selected course name 
          return_array.push(reserves_array[i][1].toString()); //Grab textbook name(s) from column 2
        }
    }
    return return_array;
  }

//Get all titles for a given barcode 
  function getTitlebyBarcode(barcode) {
    var return_array = [];
    var getLastRow = shReserves.getLastRow();
    var reserves_array = shReserves.getRange(2,1,getLastRow,3).getValues();
    for(var i = 0; i <= getLastRow-1; i++) 
    { 
      if (reserves_array[i]['2'].toString() === barcode && return_array.indexOf(reserves_array[i]['1'].toString()) === -1) { //If column 3 matches barcode in url and title is not in return array
          return_array.push(reserves_array[i][1]);
          
      }
    }
    return return_array; 
  }

// GET ITEM BARCODE AND PDF URL FROM COURSE AND TEXTBOOK 
function getBarcodeAndUrlAndId(course,textbook) {
  var getLastRow = shReserves.getLastRow() - 1; //get length of spreadsheet
  var return_array = [];
  var reserves_array = shReserves.getRange(2,1,getLastRow,5).getValues(); //get values from the first five columns of the "Reserves" spreadsheet
  for(var i = 0; i <= getLastRow - 1; i++) //Iterate through the reserves_array
    {
      if (reserves_array[i]['0'].toString() === course) { //If row one contains course name selected from dropdown 
        if(reserves_array[i]['1'].toString() === textbook) { //If row two contains textbook name selected from dropdown
            return_array.push(barcode_temp = reserves_array[i]['2'].toString()); //Grab item barcode from row 3
            return_array.push(id_temp = reserves_array[i]['3'].toString()); //Grab item ID from row 4 
            return_array.push(url_temp = reserves_array[i]['4'].toString()); //Grab PDF URL from row 5 
        }
      }
    }
    return return_array;
}

// Get item textbook, ID, and URL from barcode
function getUrlAndId(textbook) {
  var getLastRow = shReserves.getLastRow() - 1; //get length of spreadsheet
  var return_array = [];
  var reserves_array = shReserves.getRange(2,1,getLastRow,5).getValues(); //get values from the "Reserves" spreadsheet starting with the second row, column 1, and ending with the last row, column 5 

  for(var i = 0; i <= getLastRow - 1; i++) //Iterate through the reserves_array
  {
      if(reserves_array[i]['1'].toString() === textbook) { //If row one contains barcode
          return_array.push(textbook_temp = reserves_array[i][1].toString()); //Grab textbook title from row 2
          return_array.push(id_temp = reserves_array[i][3].toString()); //Grab item ID from row 4 
          return_array.push(url_temp = reserves_array[i][4].toString()); //Grab PDF URL from row 5
      }
  }
  return return_array;
}


// Get loan & expiration dates for given item barcode - updated to add textbook 7/29/22 JR
  function getLoanDate(barcode, textbook,loanLength) {
    var getLastRow = ssInUse.getLastRow() - 1; //get length of spreadsheet
    var return_array = [];
    if (getLastRow < 1) { //if In Use sheet is empty rows to get = 1
      var getLastRow = 1;
    }
    var inUse_array = ssInUse.getRange(2,1,getLastRow,3).getValues(); //get values from the first 3 columns of the "InUse" spreadsheet
    
    for(var i = 0; i <= getLastRow - 1; i++) { // Iterate through rows of the "InUse" spreadsheet 
      //REQUESTED ITEM IN USE
      if(inUse_array[i]['0'].toString() === barcode) { // If item barcode is in row one
        var loan_status = 'In Use'; // Set loan status to "In Use"
        var date_lend = inUse_array[i]['2']; //Set loan date to previous loan's exp date for the email message 
        //CUSTOM MESSAGE OPTION
        var request_status = 'Item in use. Try again later.'; //store request status message for printing 
        return_array.push(''); // Get blank lend date 
        return_array.push(date_lend); // Store expiration date in array 
        return_array.push(loan_status); // Store loan status in array
        return_array.push(request_status); //Store request status message in array
      }
    }
  
  //REQUESTED ITEM AVAILABLE 
    if (return_array.length < 1) { // If barcode not found in "InUse" spreadsheet
      var loan_status = 'Active'; // Give request status "Active" 
      //CUSTOM MESSAGE OPTION
      var request_status = 'The PDF you requested has been sent to your email. Check your email.'; //store request status message for printing 
      const now = new Date(); // Create a date object for the current date and time 
      date_exp = new Date(); // Copy current date object 
      //date_exp.setHours(date_exp.getHours() + 2); // Set expiration date-time to 2 hours later than current time 
      date_exp.setMinutes(date_exp.getMinutes() + loanLength); // Set expiration date-time to loanLength from drop down or loanDefault
      const date_lend = Utilities.formatDate(now, 'America/New_York', 'M/dd/yyyy HH:mm:ss'); // Format lend date
      date_exp = Utilities.formatDate(date_exp, 'America/New_York', 'M/dd/yyyy HH:mm:ss'); // Format expiration date 

      ssInUse.appendRow([barcode, textbook, date_exp]); //Add row to InUse spreadsheet
      return_array.push(date_lend); // Store loan date in array
      return_array.push(date_exp); // Store expiration date in array 
      return_array.push(loan_status); // Store loan status in array 
      return_array.push(request_status); //Store request status message in array

    }
    return return_array;
    }

//Add request to "Transactions spreadsheet"
function AddRecord(name, studentid, course, textbook, email, barcode, item_id, item_url, loan_status, date_lend, date_expire) { // Grab variables collected from form and functions
  if(loan_status == 'In Use') { // Do not add expiration date if there is none - 10/14/22 CT
    date_expire = '';
  }
  ssTransactions.appendRow([name, studentid, course, textbook, new Date(), email, barcode, item_id, item_url, loan_status, date_lend, date_expire]); // Add a row to "Transactions" spreadsheet with values
  
}

// GET NAME OF ITEM FROM A GIVEN BARCODE FROM THE "RESERVES" SPREADSHEET
function getTitle(itemBarcode) {
  var getLastRow = shReserves.getLastRow() - 1; //get length of spreadsheet
  var reserves_array = shReserves.getRange(2,1,getLastRow,3).getValues(); //get values from the "Reserves" spreadsheet starting with the second row, column 1, and ending with the last row, column 3
  var item_title;
  for(var i = 0; i <= getLastRow - 1; i++) //Iterate through the reserves_array
  {
      if(reserves_array[i]['2'].toString() === itemBarcode) { //If column 3 matches barcode
        item_title = reserves_array[i][1].toString(); //Grab textbook name(s) from column 2
      }
  }
  return item_title;
}

//Get course name for a given barcode from the "Reserves" spreadsheet 
function getCourseByBarcode(barcode) {
    var getLastRow = shReserves.getLastRow() - 1;
    var reserves_array = shReserves.getRange(2,1,getLastRow,3).getValues(); //get values from the "Reserves" spreadsheet starting with the second row, column 1, and ending with the last row, column 3
    var course;
    for(var i = 0; i <= getLastRow - 1; i++) //Iterate through the reserves_array
    {
      if (reserves_array[i]['2'].toString() === barcode) { //If column 3 matches barcode
        course = reserves_array[i][0].toString(); //Grab course name from column 1
      }
  }
  return course;
}

// Check for valid barcode
function checkBarcode(barcode) {
  var getLastRow = shReserves.getLastRow() - 1;
  var reserves_array = shReserves.getRange(2,1,getLastRow,3).getValues(); //get values from the "Reserves" spreadsheet starting with the second row, column 1, and ending with the last row, column 3
  var valid_barcode;
  for(var i = 0; i <= getLastRow - 1; i++) //Iterate through reserves_array
  {
      if(reserves_array[i]['2'].toString() === barcode) { //If column 3 matches barcode
        valid_barcode = 'yes';
      }
  }
  return valid_barcode;
}

// Get barcode of item from a given item_id from the "Reserves" spreadsheet
function getBarcode(item_id) { 
  var getLastRow = shReserves.getLastRow() - 1;
  var reserves_array = shReserves.getRange(2,1,getLastRow,4).getValues(); //get values from the "Reserves" spreadsheet starting with the second row, column 1, and ending with the last row, column 4
  var itemBarcode;
  for(var i = 0; i <= getLastRow - 1; i++) //Iterate through reserves array
  {
      if(reserves_array[i]['3'].toString() === item_id) { //If column 4 matches item_id
        itemBarcode = reserves_array[i]['3'].toString(); //Grab barcode from column 3
        Logger.log(itemBarcode + " ID " + item_id);
      }
  }
  return itemBarcode;  
}



//---------SHARE FILE--------
function FileShare(email, item_id, loan_status, date_expire, loanLength) {
  if(loan_status === 'Active'){
    if(loanLength > 59){//convert minutes to hours for for the email message 
      loanLength = (loanLength / 60) +' hours';
    } else {
      loanLength = loanLength +' minutes';
    }
    
    try{
      //CUSTOM MESSAGE OPTION
      var customMessage = "This PDF loan will expire in "+loanLength+". Please re-request this title if you need more time. To return this item early, please visit: ";  // Please set the custom message here.
      var resource = {role: "reader", type: "user", value: email};
      Drive.Permissions.insert(resource, item_id, {emailMessage: customMessage});
      }

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

    
    catch (e) {
      Logger.log(item_id);
      Logger.log(e);
      }
  }
    else{
      date_expire.setMinutes(date_expire.getMinutes() + 5);
      var dateExpireClean = Utilities.formatDate(date_expire, "America/New_York", "hh:mm a");
      //CUSTOM MESSAGE OPTION
      MailApp.sendEmail(email,'Requested Reserve Item Unavailable','The reserve item you requested is currently checked out and will next be available at ' + dateExpireClean + '. To re-request this item please return to our e-reserves request form: '); 
      
  
    }
  }

//---------UNSHARE FILE----------
function FileUnshare() {
  var getLastRowTransactions = ssTransactions.getLastRow() - 1;
  var getLastRowInUse = ssInUse.getLastRow() - 1;
  if (getLastRowInUse < 1) {
    var getLastRowInUse = 1;
  }
  var transactions_array = ssTransactions.getRange(2,1,getLastRowTransactions,12).getValues();
  var inuse_array = ssInUse.getRange(2,1,getLastRowInUse,3).getValues();

  for(var i = 0; i <= getLastRowTransactions - 1; i++) //Iterate through transactions_array
  {
      if(transactions_array[i][9].toString() === 'Active') { // If request is "Active"
        email = transactions_array[i][5].toString(); //Get patron email
        item_id = transactions_array[i][7].toString(); // Get item ID
        date_exp = transactions_array[i][11]; // Get expiration date
        const now = new Date(); //Create a date object for the current date and time
        
        if(date_exp < now || date_exp === now ){
          try{//if expiration date has passed unshare file
            ssTransactions.getRange(i+2,10).setValue('Expired');//change transaction status to Expired
            asset = DriveApp.getFileById(item_id) ? DriveApp.getFileById(item_id) : DriveApp.getFolderById(item_id);
            asset.removeViewer(email);
            //CUSTOM MESSAGE OPTION
            MailApp.sendEmail(email, 'Item Expired', 'Your virtual loan has expired. If you would like to borrow this document for additional time, please place a new request.')
            }
          catch (e) {
            Logger.log("File could not be unshared, please check if email address is valid");
            }
          }
        }
      }
  //remove item from in use ss if expiration date has passed
  for(var i = 0; i <= getLastRowInUse - 1; i++){
    const now = new Date();
    date_exp_in_use = inuse_array[i][2];
    if (date_exp_in_use === "") {
      continue;
      }
      else if (date_exp_in_use < now || date_exp_in_use === now) {
        ssInUse.getRange(i+2,1,1,3).clearContent();
      } 
    else {
      continue;
    }
  }
  ssInUse.sort(3); //sort InUse spreadsheet by expiration date to remove any blank rows
}
  
    


 // LIST FILES AND FOLDERS AND PUSH TO SPREADSHEET
  //UPDATE 8-11-23
  function listFilesAndFolders() {
    shReserves.clear();
        
    try {
      var parentFolder =DriveApp.getFolderById(folderOfScans);
      listFiles(parentFolder);
      listSubFolders(parentFolder,parentFolder.getName());
    } catch (e) {
      Logger.log(e.toString());
    }
    shReserves.sort(2).sort(1);
    shReserves.insertRowBefore(1)
    shReserves.getRange(1,1,1,7).setValues([["folder", "name", "description","ID", "URL", "size", "update"]]);
  }
  
  
  function listSubFolders(parentFolder,parent) {
    var childFolders = parentFolder.getFolders();//get all folders in parent folder
    while (childFolders.hasNext()) {
      var childFolder = childFolders.next();
      //Logger.log("Fold : " + childFolder.getName());
      listFiles(childFolder,parent);//get all files in child folder
      listSubFolders(childFolder,parent + "|" + childFolder.getName());
    }
  }
  
  function listFiles(fold){
    var data = [];
    var files = fold.getFiles();
    var reserves_array = [];
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
  
      reserves_array.push(data);
      //Logger.log(data);

    }
    var lastRow = shReserves.getLastRow();
    if(reserves_array.length > 0) {
    reserves_array.sort();
    shReserves.getRange(lastRow + 1,1,reserves_array.length, reserves_array[0].length).setValues(reserves_array);
  }
  
  }
  // END OF LIST FILES AND FOLDERS AND PUSH TO SPREADSHEET

