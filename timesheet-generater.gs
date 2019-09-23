/// Calender ID
var calendarId = 'xyx.com_3s0i991ci723fnu35ltu6icjdo@group.calendar.google.com';  /// this needs to be changed by you
var record_key_project = "Project:"
var record_template_project = "P1"
var record_key_workpackage = "Package:"
var record_template_workpackage = ["WP1", "WP2"]
var record_key_title = "Records:"
var max_header_columns = 30;


/**
 * Create a open translate menu item.
 * @param {Event} event The open event.
 */
function onOpen(event) {
  SpreadsheetApp.getUi()
      .createMenu('TimeSheet')
      .addItem('update', 'updateTimesheet')
      .addItem('create template', 'createTemplate')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Settings')
          .addItem('create template', 'createTemplate'))
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
 * Foramts time
 **/
function msToTime(s) {
  var ms = s % 1000;
  s = (s - ms) / 1000;
  var secs = s % 60;
  s = (s - secs) / 60;
  var mins = s % 60;
  var hrs = (s - mins) / 60;
  //return hrs + ':' + mins + ':' + secs; // milliSecs are not shown but you can use ms if needed
  return hrs + mins/60;
}



/**
 * creates the timplate sheet
 * @post update
 **/
function createTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.setColumnWidths(1, 5, 120);
  sheet.clear();
  var cell_title = sheet.getRange("C1");
  cell_title.setValue("Time Sheet")
  cell_title.setFontSize(20);
  cell_title.setHorizontalAlignment("center");
  var range_date = sheet.getRange("B2:D2");
  range_date.setValues([[new Date("8/19/2019"), "->", new Date("9/30/2019")]]);
  range_date.setHorizontalAlignment("center");
  var range_project = sheet.getRange(4,1,1,5);
  range_project.setValues([[record_key_project, record_template_project, "", "60", ""]]); 
  range_project.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  var costs = sheet.getRange("E4");
  costs.setFormulaR1C1("=R[0]C[-1]*R[0]C[-2]").setNumberFormat("#,##0.00\ [$€-1]");
  var costs_per_hour = sheet.getRange("D4");
  costs_per_hour.setNumberFormat("#,##0\ [$€/h-1]");
  
  var range_wp = sheet.getRange(6,1,1,5);
  range_wp.setValues([[record_key_workpackage, record_template_workpackage[0], "", "60", ""]]); 
  var cell_package_costs = sheet.getRange("E6");
  var costs_per_hour = sheet.getRange("D6");
  costs_per_hour.setNumberFormat("#,##0\ [$€/h-1]");
  range_wp.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  var range_wp = sheet.getRange(7,1,1,5);
  cell_package_costs.setFormulaR1C1("=R[0]C[-1]*R[0]C[-2]").setNumberFormat("#,##0.00\ [$€-1]"); 
  range_wp.setValues([[record_key_workpackage, record_template_workpackage[1], "", "60", ""]])
  var cell_package_costs = sheet.getRange("E7");
  cell_package_costs.setFormulaR1C1("=R[0]C[-1]*R[0]C[-2]").setNumberFormat("#,##0.00\ [$€-1]");
  var costs_per_hour = sheet.getRange("D7");
  costs_per_hour.setNumberFormat("#,##0\ [$€/h-1]");
  range_wp.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  var range_recording_titel = sheet.getRange(10,1);
  range_recording_titel.setValue(record_key_title)
  
  var range_recording_header = sheet.getRange(11,1,1,5);
  range_recording_header.setHorizontalAlignment("left");
  sheet.getRange(11, 5).setHorizontalAlignment("right");
  range_recording_header.setBackgroundRGB(224, 224, 224);
  range_recording_header.setValues([["Start", "Details", "", "", "Hours"]])
  sheet.getRange(11, 2, 1, 3).mergeAcross();
  range_recording_header.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * update the timesheet
 */
function updateTimesheet() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var calendar = CalendarApp.getCalendarById(calendarId)
  var record_start = new Date(sheet.getRange("B2").getValue());
  var record_end = new Date(sheet.getRange("D2").getValue())
  record_end = new Date(record_end.getTime() + 24 * 60 * 60 * 1000);
  var record_project = "x";
  var record_project_range_sum;
  for ( i = 1; i < max_header_columns; i++){
    var key = sheet.getRange(i,1).getValue();
    if(key == record_key_project){
      record_project  = sheet.getRange(i,2).getValue();
      record_project_range_sum  = sheet.getRange(i,3);
      break;
    }
  }
  var record_pkgs = {};
  var record_pkgs_range_sum = {};
  for (i = 1; i < max_header_columns; i++){
    var key = sheet.getRange(i,1).getValue();
    if(key == record_key_workpackage){
      var pgk_name  = sheet.getRange(i,2).getValue();
      record_pkgs[pgk_name] = 0;
      record_pkgs_range_sum[pgk_name] = sheet.getRange(i,3)
    }
  }
  var record_details_idx
  for (i = 1; i < max_header_columns; i++){
    var key = sheet.getRange(i,1).getValue();
    if(key == record_key_title){
      record_details_idx = i + 2;
      break;
    }
  }
  
  Logger.log(record_project)
  Logger.log(record_pkgs)
  var record_search = "#" + record_project;
  var events = calendar.getEvents(record_start, record_end,{search: record_search});
  var total_duration = 0;
  if (events.length > 0) {
    for (i = 0; i < events.length; i++) {
      var event = events[i];
      var title = event.getTitle();      
      var start = event.getStartTime() ;
      var end = event.getEndTime();
      start = new Date(start);
      end = new Date(end);
      var duration = end - start;
      var title_cuted = title.replace(record_search, '')
      for(var k in record_pkgs){
        if(title.indexOf(record_search + ":" + k) > -1) {
          record_pkgs[k] = record_pkgs[k] + duration;
          title_cuted = title_cuted.replace(":" + k, k+ " -")
        }
      }
      var splitEventId = event.getId().split('@');
      var eventURL = "https://www.google.com/calendar/event?eid=" + Utilities.base64Encode(splitEventId[0] + " " + calendarId);
      title_cuted = "=HYPERLINK(\""+ eventURL + "\";\"" + title_cuted + "\")";
      total_duration = total_duration + duration
      var range_event = sheet.getRange(record_details_idx+i, 1, 1, 5);
      sheet.getRange(record_details_idx+i, 5).setHorizontalAlignment("right");
      range_event.setValues([[start, title_cuted, " ", " ", msToTime(duration)]]) 
      sheet.getRange(record_details_idx+i,1).setNumberFormat("yyyy-MM-dd hh:mm");
      sheet.getRange(record_details_idx+i,5).setNumberFormat("0.00");
      sheet.getRange(record_details_idx+i, 2, 1, 3).mergeAcross();
      range_event.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      
      if(i % 2 != 0){
        range_event.setBackgroundRGB(248, 248, 248);
      }
    }
    record_project_range_sum.setValue(msToTime(total_duration))
    record_project_range_sum.setNumberFormat("#,##0.00\ [$h-1]");
    for(var k in record_pkgs_range_sum){
      record_pkgs_range_sum[k].setValue(msToTime(record_pkgs[k]))
      record_pkgs_range_sum[k].setNumberFormat("#,##0.00\ [$h-1]");
    }
  }
}
