//Version 0.2 Aug.20 Peiran
var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AS Front");
//var time_cell = sheet1.getRange(1, 2);
var status_cell = sheet1.getRange(2, 2);
var ui = SpreadsheetApp.getUi();

var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
var timestamp_Time = "HH:mm:ss"; 
var timestamp_Date = "yyyy-MM-dd";
var now = new Date()
var TimeStamp = Utilities.formatDate(now, timezone, timestamp_Time);
var DateStamp = Utilities.formatDate(now, timezone, timestamp_Date);

var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AS Back");
var last_row = sheet2.getLastRow();
var cell1 = sheet2.getRange(last_row + 1, 1), cell2 = sheet2.getRange(last_row + 1, 2), cell3 = sheet2.getRange(last_row + 1, 3), cell4 = sheet2.getRange(last_row, 4);
var cell5 = sheet2.getRange(last_row + 1, 5), cell6 = sheet2.getRange(last_row + 1, 6)

var dest_ID = "1HV-BzJpVv9xtEmDEKhgo5F9Uih_osLMpsFIS0psEc_k";

function BreakStart() 
  {
    var response = ui.alert('Break Start?', ui.ButtonSet.OK);
    if (response == ui.Button.OK){
        if (validation("Break Start")){    
            cell1.setValue(DateStamp);
            cell2.setValue(TimeStamp);
            cell3.setValue('Break Start');
            status_cell.setValue('On Break');
        }
    }
  }

function BreakEnd()
  {
    var response = ui.alert('Break End?', ui.ButtonSet.OK);
    if (response == ui.Button.OK){
        var last_time = sheet2.getRange(last_row, 2).getDisplayValue();
        if (validation("Break End")){
            cell1.setValue(DateStamp);
            cell2.setValue(TimeStamp);
            cell3.setValue('Break End'); 
            status_cell.setValue('Working');
            var time_diff = calculate_breaktime(last_time);
            cell5.setValue(time_diff);
            showBreakReasonsDialog();
        }
    }
  }

function ProductionStart()
  {
    var response = ui.alert('Production Start?', ui.ButtonSet.OK);
    if (response == ui.Button.OK){
        if (validation("Production Start")){
            cell1.setValue(DateStamp);
            cell2.setValue(TimeStamp);
            cell3.setValue('Production Start'); 
            status_cell.setValue('Working');
            showLateFromStartDialog();
        }
    }
  }


function ProductionFinish()
  {
    var response = ui.alert('Production Finish?', ui.ButtonSet.OK);
    if (response == ui.Button.OK){
        if (validation("Production Finish")){
            cell1.setValue(DateStamp);
            cell2.setValue(TimeStamp);
            cell3.setValue('Production Finish'); 
            status_cell.setValue('Production Finish');
            var shift_time = calculate_shifttime()[0];
            var start_index = calculate_shifttime()[1];
            cell6.setValue(shift_time);
            // Logger.log(start_index);
            push_to_SR(start_index);
            SummarizeData(shift_time);
            clear_backend();
        }
    }
  }

function showBreakReasonsDialog()
 {
  var uiDialog = HtmlService.createHtmlOutputFromFile('break_reasons_dialog').setSandboxMode(HtmlService.SandboxMode.NATIVE);
  return ui.showModalDialog(uiDialog,"Choose the break reason");
 }

 function showLateFromStartDialog()
 {
  var uiDialog = HtmlService.createHtmlOutputFromFile('late_from_start_dialog').setSandboxMode(HtmlService.SandboxMode.NATIVE);
  return ui.showModalDialog(uiDialog,"Choose the lateness");
 }

function WriteInLateness(lateness)
{
  cell4.setValue("Late From Start");
  var cell_lateness = sheet2.getRange(last_row, 5);
  cell_lateness.setValue(lateness);
}

function WriteInBreakName(breakname)
{
  cell4.setValue(breakname);
//  status_cell.setValue(breakname);
}

function validation(stampsname)
{
  var previous_cell = sheet2.getRange(last_row, 3);
  Logger.log(previous_cell.getValue());
  switch (stampsname)
  {
    case "Production Start":
      if (previous_cell.getValue() != "Production Finish" && previous_cell.getValue() != "Stamp"){
        ui.alert('Finish Shift First!');
        return false;
      } else {return true;}
      
    case "Production Finish":
      if (previous_cell.getValue() == "Break Start") {
        ui.alert('End Break First!');
        return false;
        } else if (previous_cell.getValue() == "Production Finish") {
            ui.alert('Already Finished!');
            return false;
        } else {return true;}
    
    case "Break Start":
      if (previous_cell.getValue() == "Break Start") {
          ui.alert('You are already on a Break!');
          return false;
        } else if (previous_cell.getValue() == "Production Finish"){
          ui.alert('You have not Start Production yet!');
          return false;
        } else {return true;}
    case "Break End":
        if (previous_cell.getValue() != "Break Start") {
          ui.alert('You are not on a Break!');
          return false;
        } else {return true;}
    default:
      ui.alert('Error: contact Peiran!');
  }
}

function calculate_breaktime(previous_time)
{
   var time_diff = (now.getHours() - Number(previous_time.slice(0,2))) * 60 + (now.getMinutes() - Number(previous_time.slice(3,5)));
   return time_diff;
}

function calculate_shifttime()
{
  var i = 0;
  while ((i < last_row) && (sheet2.getRange(last_row - i, 3).getDisplayValue() != "Production Start")){
    //var display_value = sheet2.getRange(last_row - i, 3).getDisplayValue();
    i++;
  }
  var shift_start = sheet2.getRange(last_row - i, 2).getDisplayValue();
  var shift_time = (now.getHours() - Number(shift_start.slice(0,2))) * 60 + (now.getMinutes() - Number(shift_start.slice(3,5)));
  return [shift_time, (last_row - i)];
}

function push_to_SR(start_index)
{
  var dest_sheet = SpreadsheetApp.openById(dest_ID).getSheetByName("Down Time Tracker Data (AS2)");
  var source_range = sheet2.getRange(start_index, 1, (last_row - start_index + 2), 7);
  var source_values = source_range.getDisplayValues();
  var target_range = dest_sheet.getRange(dest_sheet.getLastRow() + 1, 1, (last_row - start_index + 2), 7);
  target_range.setValues(source_values);
  //Summarize Timecost by reasons
}

function clear_backend()
{
  sheet2.getRange(2,1,last_row,6).clearContent();
}

function SummarizeData(shift_time)
{
  var pivot_table = SpreadsheetApp.getActive().getSheetByName("Pivot Table");
  var dest_sheet_summary = SpreadsheetApp.openById(dest_ID).getSheetByName("Submission Sheet");
  var index, reason, dest_range;
  for(index = 3; index < pivot_table.getLastRow(); index++)
  {
    reason = pivot_table.getRange(index, 1).getValue();
    switch (reason) 
    {
      case "Late From Start":
        dest_range = dest_sheet_summary.getRange(18,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break;
      case "2P/4P Changeover":
        dest_range = dest_sheet_summary.getRange(20,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break;
      case "Missing Assembly Ingredient":
        dest_range = dest_sheet_summary.getRange(21,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break;
      case "Missing Meal-kits":
        dest_range = dest_sheet_summary.getRange(22,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break;
      case "Tape Machine Down":
        dest_range = dest_sheet_summary.getRange(23,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break;         
      case "Change Tape":
        dest_range = dest_sheet_summary.getRange(24,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break;
      case "Rework":
        dest_range = dest_sheet_summary.getRange(25,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break;
      case "Other":
        dest_range = dest_sheet_summary.getRange(26,3);
        dest_range.setValue(pivot_table.getRange(index,3).getValue());
        break; 
      case "10 mins Break":
        var ten_mins_count = pivot_table.getRange(index,2).getValue();
        var ten_mins_total = pivot_table.getRange(index,3).getValue();
        var ten_mins_lateness = ten_mins_total - ten_mins_count * 10;
        break;
      case "30 mins Break":
        var thirty_mins_count = pivot_table.getRange(index,2).getValue();
        var thirty_mins_total = pivot_table.getRange(index,3).getValue();
        var thirty_mins_lateness = thirty_mins_total - thirty_mins_count * 30;
        break;
      case "Move to Kitting Line":
        //do nothing
        break;
      default:
        ui.alert('Error: contact Peiran!');
    };
    var total_lateness = ten_mins_lateness + thirty_mins_lateness;
  }
  dest_sheet_summary.getRange(12,3).setValue(shift_time);
  dest_sheet_summary.getRange(19,3).setValue(total_lateness);
}