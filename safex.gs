//safex.gs
function refresh_status() {
  //This will change a value on the SETTING page so that the status JSON formula is forced to refresh
  //update the ImportJson trigger
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings").getRange("B4").setValue(current_time()); //setting page last run time as static value needed
  //update values on status page
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C4").setValue(current_time()); //last run time
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C5").setValue(current_time_plus_n_minutes());  //next run time
  //Parse out the JSON
  cellJsonToColumns();
}

function current_time() {
  //get the current time for this user and format as universal format "yyyy-MM-dd HH:mm"
  var d = new Date();                                         // get current date/time of user
  var t = "GMT-" + d.getTimezoneOffset()/60;                  // get timezone offset (TODO: formulate flipping GMT sign)
  var date = Utilities.formatDate(d, t, "yyyy-MM-dd HH:mm");  // format date as we want (full format is: "yyyy-MM-dd'T'HH:mm:ss'Z'")
 return date;                                                 // return the result 
}

function current_time_plus_n_minutes() {
//add n minutes to a date value (from Status page)
  var m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C6").getValue()   // m=minutes to add
  var d = new Date(); d.setMinutes(d.getMinutes() + m);                                              //d = current date/time
  var t = "GMT-" + d.getTimezoneOffset()/60                                                          // get timezone offset (TODO: formulate flipping GMT sign)
  date = Utilities.formatDate(d, t, "yyyy-MM-dd HH:mm"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"               // format date as we want (full format is: "yyyy-MM-dd'T'HH:mm:ss'Z'")
 return date;                                                                                        // return the result 
}

function charToAsc (inp) {
  //the description data is encoded in unicode decimal format. 
  //we need to convert the decimal code to ascii character
  var d = "";                                     //set up placeholder variable
  var e = "";                                     //set up placeholder variable
  var b = inp.split(',').map(function(item) {     //split the description on commas
    d = String.fromCharCode(parseInt(item, 10));  //convert decimal to letter
    e = e+d;                                      //concatenate the letters to make a full string
  });
  //console.log (e);                              //used for debug
  return e;                                       //return the result
}

function cellJsonToColumns () {
 //SUMMARY: take JSON from a cell and get one parse it across columns

 //10. get the sheet to process
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("offers");
 
 //20.get current used range (assume A1 is start)
  var first_row = 2;          //we expect data in row 2 after data is imported 
  var last_row = sheet.getLastRow();    //we find the last row used
  var last_col = sheet.getLastColumn(); //we find the last column used
//unused//  var rng = sheet +"!"+ sheet.getDataRange().getA1Notation();

 //30. get current column headers
//unused//  var header = sheet.getRange(1, 1, 1, last_col).getValues()[0];
 
 //40. get column with the JSON data
  var src_col = 2;   //the JSON is column B(2), although unicoded as decimals
  var src_cell="";   //placeholder for our loop

 //50.set up variables to use
  var json;         //used later to assogn JSOn to it
  var keys = [];    //keys of the JSON
  var newCols = [];      //placeholder for new columns as needed
  var newValsRow = [];   //placeholder for new row values
  var newValsSheet = []; //placeholder for new sheet values
  
 //60.DEBUG override values section for testing
    //first_row = 251    //hardcode a specific row to start
    //last_row = 251;   //hardcode a specific row to stop
 
 //70. calcualte number of loops
  var loopCount = (last_row - first_row + 1); //calculate the number of loops we need

 //80. loop each row (r)
 for (var r = first_row; r <= last_row; r++) {
    
    //80.20 Get the raw data (will be a string of comma seprated decimals in this case)
    src_cell = sheet.getRange(r,src_col).getValue();

    //80.30 Convert using custom function from decimals to letters
    src_cell = charToAsc(src_cell);    

    //80.40 if src_cell like descrription then continue
      if(src_cell.includes("twm_version")) {
      
      //80.40.10 make the string variable into JSON
      json = JSON.parse(src_cell);      // parse each json cell

      //80.40.20 get the keys from the JSON (can be used to compare to column headers)
      keys = Object.keys(json);         // parse all keys
        
      //80.40.30 add each of the key values to an array for the row
      newValsRow = [src_cell,json["twm_version"],json["description"],json["main_image"]
                   ,json["image_2"],json["image_3"],json["image_4"],json["sku"]
                   ,json["barcode"],json["country"],json["shipping"],json["nft"]
                   ,json["open_message"]];

     //80.40.40 add each of these to an array for the page
      newValsSheet.push(newValsRow);

   } else {newValsRow = [src_cell,"",src_cell,"","","","","","","","","",""];
      newValsSheet.push(newValsRow);} //end 80.40 if evaluation

 } // end 80. loop of rows

 //90. then update range at one time uses plural setValues along with the array
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("offers").getRange("R" + first_row + ":AD" + last_row).setValues(newValsSheet);
  
 //100. write to console when we are done
 // console.log ("rows updated: " + last_row);
}
