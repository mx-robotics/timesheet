/**
 * TimeSheet
 * @url https://github.com/mx-robotics/timesheet
 * @brief This script extracts calendar entries to generate a time sheet
 * @installation
 *  - create a new sheet
 *  - under "Tools" -> "Script Editor" copy and pased this stript
 *  - get your calendarId: check your calendar properties and serarch your for your calendar id and past it into the config sheet.
 *  - close and reopen the new sheet you will be asked for access rights which you have to give.
 *  - have fun
 **/

function date_next_recort(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getActiveSheet();
  var timezone = ss.getSpreadsheetTimeZone();
  var date = new Date(sheet.getRange("B1").getValue());
  var record_next = new Date(date);
  var month = record_next.getMonth();
  record_next.setMonth(month + 1);
  Logger.log("ScriptTimeZone: " + Session.getScriptTimeZone());
  Logger.log("SheetTimeZone:  " + ss.getSpreadsheetTimeZone()); 
  Logger.log('record_start: ' + Utilities.formatDate(date, timezone, 'MMMM dd, yyyy'));
  Logger.log('record_next:  ' + Utilities.formatDate(record_next, timezone, 'MMMM dd, yyyy'));
  return record_next;
}

// Some global variables
var calendarId;
var project_prefix;
var sheet_cell_date;
var sheet_row_first_day;
var sheet_col_day;
var sheet_row_workpackages;
var sheet_col_first_workpackages;
var sheet_col_last_workpackages;

/**
 * Reads a Value by a key form a 2D list
 * @param cells list
 * @param variable_name
 */
function getValueByKeyFromList2D(list2D, variable_name){
  for (var r = 0; r < list2D.length; r++){
    var key = list2D[r][0];
    if (key == variable_name){
      var value = list2D[r][1];
      return value;
    }
  }
  Browser.msgBox("Config \nVariable: \'" + variable_name + "\' not found");
  exit();
}

/**
 * updates all config variables
 */
function updateConfig(){
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var list2D = sheet.getRange(4,1,30,2 ).getValues()
  calendarId = getValueByKeyFromList2D(list2D,'calendarId');
  project_prefix = getValueByKeyFromList2D(list2D,'project_prefix');
  sheet_cell_date = getValueByKeyFromList2D(list2D,'sheet_cell_date');
  sheet_row_first_day = getValueByKeyFromList2D(list2D,'sheet_row_first_day');
  sheet_col_days = getValueByKeyFromList2D(list2D,'sheet_col_days');
  sheet_row_workpackages = getValueByKeyFromList2D(list2D,'sheet_row_workpackages');
  sheet_col_first_workpackages =getValueByKeyFromList2D(list2D,'sheet_col_first_workpackages');
  sheet_col_last_workpackages = getValueByKeyFromList2D(list2D,'sheet_col_last_workpackages');
  sheet_wkp_name_error = getValueByKeyFromList2D(list2D,'sheet_wkp_name_error');
  sheet_wkp_name_note = getValueByKeyFromList2D(list2D,'sheet_wkp_name_note');
  sheet_wkp_name_description = getValueByKeyFromList2D(list2D,'sheet_wkp_name_description');
}

/**
 * Helper to check if a fild holds a date variable
 **/
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return true;
}

/**
 * Foramts time
 **/
function msToTime(s) {
  var ms = s % 1000;
  s = (s - ms) / 1000;
  var secs = s % 60;
  s = (s - secs) / 60;
  var mins = s % 60;
  var hrs = (s - mins) / 60;;
  return hrs + ':' + Utilities.formatString("%02d", mins); // milliSecs are not shown but you can use ms if needed
}


/**
 * Create a open translate menu item.
 * @param {Event} event The open event.
 */
function onOpen(event) {
  updateConfig()
  SpreadsheetApp.getUi()
      .createMenu('TimeSheet')
      .addItem('update Config', 'updateConfig')
      .addItem('update Current Row', 'updateCurrentRow')
      .addItem('update Current Sheet', 'updateCurrentSheet')
      .addToUi();
}

/**
 * Open the Add-on upon install.
 * @param {Event} event The install event.
 */
function onInstall(event) {
  onOpen(event);
}

/**
 * Returns a list with all workpackage entries and the realted columns  
 * @param sheet 
 * @return list with [workpackage, column]
 */
function getWorkpackages(sheet) {
  var wkps = [];
  for(var c = 1; c <= sheet_col_last_workpackages; c++){
    if(sheet.getRange(sheet_row_workpackages,c).isBlank()){
      continue;
    }    
    workpkg = sheet.getRange(sheet_row_workpackages,c).getValue();
    wkps.push([workpkg, c])
    if(workpkg.toLowerCase() == sheet_wkp_name_error.toLowerCase()){
      return wkps;
    }
  }
  throw new Error("the last workpackage must be the error package")

}
/**
 * Returns the column index with the description
 * @return column idx 
 */
function getColDescription(sheet) {
  for(var c = 1; c <= sheet_col_last_workpackages; c++){
    if(sheet.getRange(sheet_row_workpackages,c).isBlank()){
      continue;
    }    
    workpkg = sheet.getRange(sheet_row_workpackages,c).getValue();
    if(workpkg.toLowerCase() == sheet_wkp_name_description.toLowerCase()){
      return c;
    }
  }
  throw new Error("No description title in workpackage row")
}

/**
 * Returns the date of a row
 * @return date 
 */
function getDateFromRow(sheet, row){
  sheet.getRange(row,1).setBackground("red");
  var range_date = sheet.getRange(row,sheet_col_days);
  if( !isValidDate(range_date.getValue()) ) {
    Browser.msgBox('Could not read the date off row ' + row);
    throw new Error("date not readable")
  }
  var date = new Date(range_date.getValue());
  return date;
}

/**
 * updates the work package durations on a spezific row
 * @return sheet 
 * @return workpkgs 
 * @return row 
 * @return col_des 
 */
function updateRow(sheet, workpkgs, row, col_des){
  var date = getDateFromRow(sheet, row)  
  var record_search = "#" + project_prefix;
  calendar = CalendarApp.getCalendarById(calendarId);
  var end_date = new Date(date);
  end_date.setDate(end_date.getDate() + 1)
  var events = calendar.getEvents(date, end_date, {search: record_search});
  var description = String();
  var duration_error = 0;
  for (i = 0; i < events.length; i++) {
    event = events[i];   
    duration_error = duration_error + (event.getEndTime() - event.getStartTime());
  }
  var duration_wkp = [];
  for(j = 0; j < workpkgs.length; j++){
    duration_wkp.push(0)
    var wkp = workpkgs[j][0];
    var wkp_search = "#" + project_prefix + ":" + wkp;
    for (i = 0; i < events.length; i++) {
      event = events[i];
      var title = event.getTitle(); 
      var idx = title.toLowerCase().includes(wkp_search.toLowerCase())
      if(idx > 0 ){
        var duration = event.getEndTime() - event.getStartTime()
        duration_wkp[j] = duration_wkp[j] + duration;
        duration_error = duration_error - duration
        var title_cuted = title.replace(wkp_search, '').trim();
        if(title_cuted.length > 1){
          /// push descriptions to check for double entries
          title_cuted = wkp + ": " + title_cuted;
          if(description.length == 0){          
            description = title_cuted;
          }else {
            description = description + ", " + title_cuted;
          }
        }
      }
    }
    if(wkp.toLowerCase() == sheet_wkp_name_error.toLowerCase()){
      duration_wkp[j] = duration_error
    }    
  }
  for(j = 0; j < workpkgs.length; j++){
    if(duration_wkp[j] == 0) duration_wkp[j] = '';
    else duration_wkp[j] = msToTime(duration_wkp[j]);
  }
  range = sheet.getRange(row,workpkgs[0][1], 1, workpkgs[workpkgs.length-1][1] - workpkgs[0][1] + 1);
  range.setValues([duration_wkp])
  range = sheet.getRange(row, col_des).setValue(description)
  //Logger.log(duration_wkp)
  sheet.getRange(row,1).setBackground("white");
}

/**
 * updates the current row
 */
function updateCurrentRow(){
  updateConfig();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  workpkgs = getWorkpackages(sheet);
  col_des = getColDescription(sheet);
  var active_row_idx = sheet.getActiveRange().getRowIndex();
  updateRow(sheet, workpkgs, active_row_idx, col_des);
}

/**
 * updates the current sheet
 */
function updateCurrentSheet(){
  updateConfig();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  workpkgs = getWorkpackages(sheet);
  col_des = getColDescription(sheet);
  var row = sheet_row_first_day;
  var loop = true;
  while(loop){
    var range_date = sheet.getRange(row,sheet_col_days);
    if( isValidDate(range_date.getValue()) ) {
      updateRow(sheet, workpkgs, row++, col_des);
    } else {
      loop = false;
    }

  }
}
