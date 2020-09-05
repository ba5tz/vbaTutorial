function onEdit(e) {
  
  var ss = SpreadsheetApp.getActiveSheet();
  
  var rng = e.range;
  
  if (rng.getColumn() == 2 && rng.getValue() != ''){
  var cell = ss.getRange(rng.getRow(), 1);
    var waktu = Utilities.formatDate( new Date(),'GMT+7', 'dd MMM yyyy hh:mm:ss');
    cell.setValue(waktu)  //timestamp
    
    var cfcell = ss.getRange(rng.getRow(), 1, 1, 5);
    cfcell.setBorder(true, true, true, true, true, false, 'red',SpreadsheetApp.BorderStyle.DOTTED);
  }else if (rng.getValue() == ''){
   var cell = ss.getRange(rng.getRow(), 1);
    cell.setValue('')  //timestamp
    
    var cfcell = ss.getRange(rng.getRow(), 1, 1, 5);
    cfcell.setBorder(true, false, false, false, false, false, 'red',SpreadsheetApp.BorderStyle.DOTTED);
  
  }
}
