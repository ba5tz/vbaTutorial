function onEdit(e) {
  
  var namasheet = 'Tagian';
  var kolom = '4';
  var nilai = 'Lunas';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var rng = e.range;

  if (namasheet == ss.getSheetName() && kolom == rng.getColumn()) {
    if ( nilai == rng.getValue()){
      sheet.hideRows(rng.getRow());
    }
  }
}
