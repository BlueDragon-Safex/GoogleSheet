function refresh_status() {
  //This will change a value on the SETTING page so that the status JSON formula is forced to refresh
  //set trigger
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings").getRange("B4").setValue(current_time());
 //update status page
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C4").setValue(current_time());
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C5").setValue(add_minutes());
}

function current_time() {
  //get the current time and format as universal format
  var d = new Date();
  var t = "GMT-" + d.getTimezoneOffset()/60;
  var date = Utilities.formatDate(d, t, "yyyy-MM-dd HH:mm"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"
 return date;
}

function add_minutes() {
//add minutes to a data value
  var m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status").getRange("C6").getValue()
  var d = new Date(); d.setMinutes(d.getMinutes() + m);
  var t = "GMT-" + d.getTimezoneOffset()/60
  date = Utilities.formatDate(d, t, "yyyy-MM-dd HH:mm"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"
 return date;
}
