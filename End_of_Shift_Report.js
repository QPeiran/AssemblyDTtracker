//Submiting the end of shift report every shift
function submit()
{
  const ui = SpreadsheetApp.getUi();
  
  var response = ui.alert('Do you want to submit?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) 
  {
    const front_sheet = SpreadsheetApp.getActive().getSheetByName("Submission Sheet");
    var range_to_copy = front_sheet.getRange(110,1,3,47);
    var destination_sheet = SpreadsheetApp.getActive().getSheetByName("Historical Submissions");
    var destination = destination_sheet.getRange(destination_sheet.getLastRow() + 1, 1);
    range_to_copy.copyTo(destination,{contentsOnly:true});
    var range_to_clear = front_sheet.getRangeList(['B3:C6', 'B9:C10', 'B12:C13', 'B18:C25', 'B33:C35', 'B51:B52', 'B55:C56', 'B63:D65']);
    range_to_clear.clearContent();
    ui.alert('Submitted!');
  } 
  else 
  {ui.alert('Not Submitted!');}

  // TODO: submit via email
  //
}
