function createObjectsFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet 1'); 
  var columnB = sheet.getRange('B:B').getValues(); 
  var TheTable = [['Cell', 'Date', 'Candidate_id', 'Position_id']]; // Add header row

  for (var i = 0; i < columnB.length; i++) {
    if (columnB[i][0]) { 
      var cellName = 'B' + (i+1).toString();
      
      var cellContent = columnB[i][0].toString(); 
      var datePattern = /[A-Z][a-z]{2} [A-Z][a-z]{2}\. \d{1,2}, \d{4}, \d{1,2}:\d{2} [A|P]M UTC/g;
      var dateMatches = cellContent.match(datePattern); 

      var dateString = '';
      if (dateMatches) {
        var dateObj = new Date(dateMatches[0]);
        dateString = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
      }

      var idPattern = /\b[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}\b/g; 
      var idMatches = cellContent.match(idPattern);
      
      if (idMatches) {
        for (var j = 0; j < idMatches.length; j += 2) {
          var pair = [
            cellName,
            dateString,
            idMatches[j],
            idMatches[j+1] ? idMatches[j+1] : ''
          ];
          TheTable.push(pair);
        }
      }
    }
  }

  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Output2'); 
  outputSheet.getRange(1, 1, TheTable.length, 4).setValues(TheTable);
}
