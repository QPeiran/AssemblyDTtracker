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
            showDialog();
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
            var shift_time = calculate_shifttime();
            cell6.setValue(shift_time);
        }
    }
  }

function showDialog()
 {
  var uiDialog = HtmlService.createHtmlOutputFromFile('ASBreakReasons').setSandboxMode(HtmlService.SandboxMode.NATIVE);
  return SpreadsheetApp.getUi().showModalDialog(uiDialog,"Choose the break reason");
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
        SpreadsheetApp.getUi().alert('Finish Shift First!');
        return false;
      } else {return true;}
      
    case "Production Finish":
      if (previous_cell.getValue() == "Break Start") {
        SpreadsheetApp.getUi().alert('End Break First!');
        return false;
        } else if (previous_cell.getValue() == "Production Finish") {
            SpreadsheetApp.getUi().alert('Already Finished!');
            return false;
        } else {return true;}
    
    case "Break Start":
      if (previous_cell.getValue() == "Break Start") {
          SpreadsheetApp.getUi().alert('You are already on a Break!');
          return false;
        } else if (previous_cell.getValue() == "Production Finish"){
          SpreadsheetApp.getUi().alert('You have not Start Production yet!');
          return false;
        } else {return true;}
    case "Break End":
        if (previous_cell.getValue() != "Break Start") {
          SpreadsheetApp.getUi().alert('You are not on a Break!');
          return false;
        } else {return true;}
    default:
      SpreadsheetApp.getUi().alert('Error: contact Peiran!');
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
  return shift_time;
}

function push_to_SR()
{
  var dest_rep = ""
  //TODO
  
}