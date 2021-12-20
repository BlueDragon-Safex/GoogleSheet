function refresh_status() {
  //This will change a value on the SETTING page so that the status JSON formula is forced to refresh
  //update the ImportJson trigger
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings").getRange("B4").setValue(current_time()); //setting page last run time as static value needed
  //update values on status page
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C4").setValue(current_time()); //last run time
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C5").setValue(current_time_plus_n_minutes());  //next run time
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
