/**
 * TimeSheet (optimized)
 * @brief This script extracts calendar entries to generate a time sheet
 * @installation
 *  - create a new sheet
 *  - under "Tools" -> "Script Editor" copy and paste this script
 *  - get your calendarId: check your calendar properties, search for your calendar id and paste it into the config sheet.
 *  - close and reopen the new sheet; you will be asked for access rights which you have to give.
 *  - have fun
 *
 * Performance notes vs. original:
 *  - The calendar is fetched ONCE and the entire month's events are queried in a
 *    single getEvents() call, then bucketed by day in memory. The original made
 *    one getCalendarById() + getEvents() call PER ROW (≈30 round-trips per month).
 *  - Config / workpackage / description rows are read with a single getValues()
 *    instead of cell-by-cell getValue()/isBlank() calls.
 *  - All results for the whole sheet are collected in memory and written with a
 *    single setValues() call, instead of one write (+ red/white background flip)
 *    per row.
 **/

// Some global variables
var calendarId;
var project_prefix;
var sheet_cell_date;
var sheet_row_first_day;
var sheet_col_days;
var sheet_row_workpackages;
var sheet_col_first_workpackages;
var sheet_col_last_workpackages;
var sheet_wkp_name_error;
var sheet_wkp_name_note;
var sheet_wkp_name_description;

/**
 * Reads a value by a key from a 2D list
 */
function getValueByKeyFromList2D(list2D, variable_name){
  for (var r = 0; r < list2D.length; r++){
    if (list2D[r][0] == variable_name){
      return list2D[r][1];
    }
  }
  Browser.msgBox("Config \nVariable: \'" + variable_name + "\' not found");
  throw new Error("Config variable not found: " + variable_name);
}

/**
 * updates all config variables
 */
function updateConfig(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var list2D = sheet.getRange(4,1,30,2).getValues();
  calendarId = getValueByKeyFromList2D(list2D,'calendarId');
  project_prefix = getValueByKeyFromList2D(list2D,'project_prefix');
  sheet_cell_date = getValueByKeyFromList2D(list2D,'sheet_cell_date');
  sheet_row_first_day = getValueByKeyFromList2D(list2D,'sheet_row_first_day');
  sheet_col_days = getValueByKeyFromList2D(list2D,'sheet_col_days');
  sheet_row_workpackages = getValueByKeyFromList2D(list2D,'sheet_row_workpackages');
  sheet_col_first_workpackages = getValueByKeyFromList2D(list2D,'sheet_col_first_workpackages');
  sheet_col_last_workpackages = getValueByKeyFromList2D(list2D,'sheet_col_last_workpackages');
  sheet_wkp_name_error = getValueByKeyFromList2D(list2D,'sheet_wkp_name_error');
  sheet_wkp_name_note = getValueByKeyFromList2D(list2D,'sheet_wkp_name_note');
  sheet_wkp_name_description = getValueByKeyFromList2D(list2D,'sheet_wkp_name_description');
}

/**
 * Helper to check if a field holds a date variable
 **/
function isValidDate(d) {
  return Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d.getTime());
}

/**
 * Formats time
 **/
function msToTime(s) {
  var ms = s % 1000;
  s = (s - ms) / 1000;
  var secs = s % 60;
  s = (s - secs) / 60;
  var mins = s % 60;
  var hrs = (s - mins) / 60;
  return hrs + ':' + Utilities.formatString("%02d", mins);
}

/**
 * Local-date key (yyyy-mm-dd) used to bucket events by day.
 */
function dayKey(d){
  return d.getFullYear() + '-' + (d.getMonth()+1) + '-' + d.getDate();
}

/**
 * Create the menu.
 */
function onOpen(event) {
  updateConfig();
  SpreadsheetApp.getUi()
      .createMenu('TimeSheet')
      .addItem('update Config', 'updateConfig')
      .addItem('update Current Row', 'updateCurrentRow')
      .addItem('update Current Sheet', 'updateCurrentSheet')
      .addToUi();
}

function onInstall(event) {
  onOpen(event);
}

/**
 * Reads the workpackage row once and returns [workpackage, column] pairs,
 * stopping at the error package. Single getValues() call.
 */
function getWorkpackages(sheet) {
  var lastCol = sheet_col_last_workpackages;
  var rowVals = sheet.getRange(sheet_row_workpackages, 1, 1, lastCol).getValues()[0];
  var wkps = [];
  for (var c = 0; c < lastCol; c++){
    var workpkg = rowVals[c];
    if (workpkg === '' || workpkg === null) continue;
    wkps.push([workpkg, c + 1]); // store 1-based column
    if (String(workpkg).toLowerCase() == String(sheet_wkp_name_error).toLowerCase()){
      return wkps;
    }
  }
  throw new Error("the last workpackage must be the error package");
}

/**
 * Returns the 1-based column index with the description. Single getValues() call.
 */
function getColDescription(sheet) {
  var lastCol = sheet_col_last_workpackages;
  var rowVals = sheet.getRange(sheet_row_workpackages, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < lastCol; c++){
    var workpkg = rowVals[c];
    if (workpkg === '' || workpkg === null) continue;
    if (String(workpkg).toLowerCase() == String(sheet_wkp_name_description).toLowerCase()){
      return c + 1;
    }
  }
  throw new Error("No description title in workpackage row");
}

/**
 * Computes [durations[], description] for the events of a single day.
 * Pure in-memory; no API calls.
 */
function computeRow(dayEvents, workpkgs){
  var description = '';
  var duration_error = 0;
  for (var i = 0; i < dayEvents.length; i++) {
    duration_error += (dayEvents[i].getEndTime() - dayEvents[i].getStartTime());
  }
  var duration_wkp = [];
  for (var j = 0; j < workpkgs.length; j++){
    duration_wkp.push(0);
    var wkp = workpkgs[j][0];
    var wkp_search = ("#" + project_prefix + ":" + wkp).toLowerCase();
    for (var k = 0; k < dayEvents.length; k++) {
      var title = dayEvents[k].getTitle();
      if (title.toLowerCase().includes(wkp_search)) {
        var duration = dayEvents[k].getEndTime() - dayEvents[k].getStartTime();
        duration_wkp[j] += duration;
        duration_error -= duration;
        var title_cuted = title.replace(new RegExp(wkp_search, 'i'), '').trim();
        if (title_cuted.length > 1){
          title_cuted = wkp + ": " + title_cuted;
          description = description.length === 0 ? title_cuted : description + ", " + title_cuted;
        }
      }
    }
    if (String(wkp).toLowerCase() == String(sheet_wkp_name_error).toLowerCase()){
      duration_wkp[j] = duration_error;
    }
  }
  for (var m = 0; m < duration_wkp.length; m++){
    duration_wkp[m] = duration_wkp[m] == 0 ? '' : msToTime(duration_wkp[m]);
  }
  return [duration_wkp, description];
}

/**
 * updates the current row
 */
function updateCurrentRow(){
  updateConfig();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var workpkgs = getWorkpackages(sheet);
  var col_des = getColDescription(sheet);
  var row = sheet.getActiveRange().getRowIndex();

  var dateVal = sheet.getRange(row, sheet_col_days).getValue();
  if (!isValidDate(dateVal)) {
    Browser.msgBox('Could not read the date off row ' + row);
    throw new Error("date not readable");
  }
  var date = new Date(dateVal);
  var end_date = new Date(date);
  end_date.setDate(end_date.getDate() + 1);

  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(date, end_date, {search: "#" + project_prefix});

  var res = computeRow(events, workpkgs);
  var firstCol = workpkgs[0][1];
  var width = workpkgs[workpkgs.length-1][1] - firstCol + 1;
  sheet.getRange(row, firstCol, 1, width).setValues([res[0]]);
  sheet.getRange(row, col_des).setValue(res[1]);
}

/**
 * updates the current sheet.
 * Reads all dates at once, queries the whole date range in ONE getEvents call,
 * buckets events by day, then writes all rows in batched setValues calls.
 */
function updateCurrentSheet(){
  updateConfig();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var workpkgs = getWorkpackages(sheet);
  var col_des = getColDescription(sheet);

  // 1. Read all day rows in one shot, find the contiguous block of valid dates.
  var firstRow = sheet_row_first_day;
  var maxRows = sheet.getMaxRows() - firstRow + 1;
  if (maxRows <= 0) return;
  var dateCol = sheet.getRange(firstRow, sheet_col_days, maxRows, 1).getValues();

  var rows = [];          // [{row, date}]
  for (var r = 0; r < dateCol.length; r++){
    var v = dateCol[r][0];
    if (!isValidDate(v)) break; // stop at first non-date, matching original behaviour
    rows.push({ row: firstRow + r, date: new Date(v) });
  }
  if (rows.length === 0) return;

  // 2. One calendar query spanning the full range.
  var rangeStart = new Date(rows[0].date);
  var rangeEnd = new Date(rows[rows.length-1].date);
  rangeEnd.setDate(rangeEnd.getDate() + 1);

  var calendar = CalendarApp.getCalendarById(calendarId);
  var allEvents = calendar.getEvents(rangeStart, rangeEnd, {search: "#" + project_prefix});

  // 3. Bucket events by local day.
  var buckets = {};
  for (var e = 0; e < allEvents.length; e++){
    var key = dayKey(allEvents[e].getStartTime());
    (buckets[key] || (buckets[key] = [])).push(allEvents[e]);
  }

  // 4. Compute every row in memory.
  var firstCol = workpkgs[0][1];
  var width = workpkgs[workpkgs.length-1][1] - firstCol + 1;
  var durBlock = [];   // 2D array for the workpackage columns
  var desBlock = [];   // 2D array (single column) for descriptions
  for (var i = 0; i < rows.length; i++){
    var dayEvents = buckets[dayKey(rows[i].date)] || [];
    var res = computeRow(dayEvents, workpkgs);
    durBlock.push(res[0]);
    desBlock.push([res[1]]);
  }

  // 5. Two batched writes for the entire sheet.
  sheet.getRange(rows[0].row, firstCol, durBlock.length, width).setValues(durBlock);
  sheet.getRange(rows[0].row, col_des, desBlock.length, 1).setValues(desBlock);
}
